import os
import io
import json
import sqlite3
import functools
import pandas as pd

# Load .env file for local development (ignored in production where env vars are set directly)
_env_path = os.path.join(os.path.dirname(__file__), '.env')
if os.path.exists(_env_path):
    with open(_env_path) as _f:
        for _line in _f:
            _line = _line.strip()
            if _line and not _line.startswith('#') and '=' in _line:
                _k, _v = _line.split('=', 1)
                os.environ.setdefault(_k.strip(), _v.strip())
from flask import Flask, render_template, request, redirect, url_for, flash, jsonify, send_file, g, session
from werkzeug.security import generate_password_hash, check_password_hash
from datetime import datetime, timedelta

try:
    import outlook_connector as _oc
    _OUTLOOK_AVAILABLE = True
except Exception:
    _oc = None
    _OUTLOOK_AVAILABLE = False

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'cpainc2026')

# In production (Railway), DATA_DIR points to the mounted persistent volume.
# Locally it defaults to the project folder.
_DATA_DIR = os.environ.get('DATA_DIR', os.path.dirname(__file__))
os.makedirs(_DATA_DIR, exist_ok=True)
DATABASE = os.path.join(_DATA_DIR, 'CPAinc.sqlite')

# ── Email configuration (set these as Railway environment variables) ──────────
MAIL_SERVER   = os.environ.get('MAIL_SERVER',   'smtp.gmail.com')
MAIL_PORT     = int(os.environ.get('MAIL_PORT', '587'))
MAIL_USERNAME = os.environ.get('MAIL_USERNAME', '')
MAIL_PASSWORD = os.environ.get('MAIL_PASSWORD', '')
MAIL_FROM     = os.environ.get('MAIL_FROM',     MAIL_USERNAME)
MAIL_TO       = os.environ.get('MAIL_TO',       '')   # comma-separated alert recipients
DAILY_TASK_KEY= os.environ.get('DAILY_TASK_KEY','')   # secret key for /admin/run-daily-tasks

SESSION_TIMEOUT_SECONDS = 3600  # 1 hour


def send_email(to_addrs, subject, body_html, attachments=None):
    """Send an email via SMTP. to_addrs can be a string or list. Returns (ok, error)."""
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.text      import MIMEText
    from email.mime.base      import MIMEBase
    from email                import encoders as _enc
    if not MAIL_USERNAME or not MAIL_PASSWORD:
        return False, 'Email not configured (MAIL_USERNAME / MAIL_PASSWORD missing)'
    if isinstance(to_addrs, str):
        to_addrs = [a.strip() for a in to_addrs.split(',') if a.strip()]
    if not to_addrs:
        return False, 'No recipients'
    try:
        msg = MIMEMultipart('mixed')
        msg['From']    = MAIL_FROM or MAIL_USERNAME
        msg['To']      = ', '.join(to_addrs)
        msg['Subject'] = subject
        msg.attach(MIMEText(body_html, 'html'))
        if attachments:
            for fname, data in attachments:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(data)
                _enc.encode_base64(part)
                part.add_header('Content-Disposition', f'attachment; filename="{fname}"')
                msg.attach(part)
        with smtplib.SMTP(MAIL_SERVER, MAIL_PORT, timeout=15) as s:
            s.ehlo()
            s.starttls()
            s.login(MAIL_USERNAME, MAIL_PASSWORD)
            s.sendmail(msg['From'], to_addrs, msg.as_string())
        return True, None
    except Exception as e:
        return False, str(e)

def get_db():
    if 'db' not in g:
        g.db = sqlite3.connect(DATABASE)
        g.db.row_factory = sqlite3.Row
    return g.db

@app.teardown_appcontext
def close_db(e=None):
    db = g.pop('db', None)
    if db:
        db.close()

# ── Filters ───────────────────────────────────────────────────────────────────

@app.template_filter('fromjson')
def fromjson_filter(value):
    if not value:
        return {}
    try:
        return json.loads(value)
    except Exception:
        return {}

@app.template_filter('us_date')
def us_date_filter(value):
    """Convert YYYY-MM-DD (or datetime) to the configured display format."""
    if not value:
        return '—'
    try:
        s = str(value)[:10]
        return datetime.strptime(s, '%Y-%m-%d').strftime(get_date_format())
    except Exception:
        return value

@app.template_filter('fmtdate')
def _to_iso(raw):
    """Convert any pipeline date value to a YYYY-MM-DD string, or return None.
    Handles ISO format (2026-10-06 or 2026-10-06T00:00:00.000Z)
    and US format (10/6/2026 or 10/23/2026) as stored by some Railway imports."""
    if not raw:
        return None
    s = str(raw).strip()
    if not s:
        return None
    # ISO: starts with 4-digit year
    if len(s) >= 10 and s[4] == '-':
        return s[:10]
    # US format MM/DD/YYYY or M/D/YYYY
    try:
        return datetime.strptime(s.split('T')[0], '%m/%d/%Y').strftime('%Y-%m-%d')
    except Exception:
        pass
    return None


def fmt_date(val):
    if not val:
        return ''
    try:
        s = str(val)[:10]
        return datetime.strptime(s, '%Y-%m-%d').strftime(get_date_format())
    except Exception:
        return str(val)[:10]

@app.template_filter('fmtcurrency')
def fmt_currency(val):
    if val is None or val == '':
        return ''
    try:
        return f'${float(val):,.2f}'
    except Exception:
        return str(val)

@app.template_filter('fmtpct')
def fmt_pct(val):
    if val is None or val == '':
        return ''
    try:
        return f'{float(val)*100:.1f}%'
    except Exception:
        return str(val)

@app.template_filter('format_number')
def format_number(val):
    try:
        return f'{int(val):,}'
    except Exception:
        return str(val) if val is not None else '—'

# ── Settings helpers ──────────────────────────────────────────────────────────

COMMISSION_SPLIT = 0.60

def get_commission_split():
    try:
        row = get_db().execute('SELECT value FROM Settings WHERE key = "commission_split"').fetchone()
        return float(row[0]) if row else COMMISSION_SPLIT
    except Exception:
        return COMMISSION_SPLIT

def get_payment_tolerance():
    try:
        row = get_db().execute('SELECT value FROM Settings WHERE key = "payment_tolerance"').fetchone()
        return float(row[0]) if row else 0.01
    except Exception:
        return 0.01

def get_account_splits():
    try:
        rows = get_db().execute('SELECT account_name, countries, split_rate FROM AccountSplits').fetchall()
        result = []
        for r in rows:
            countries = set(c.strip().upper() for c in (r[1] or '').split(',') if c.strip())
            result.append((r[0].lower(), countries, float(r[2])))
        return result
    except Exception:
        return []

def get_kristin_split():
    try:
        row = get_db().execute('SELECT value FROM Settings WHERE key = "kristin_split"').fetchone()
        return float(row[0]) if row else 0.70
    except Exception:
        return 0.70

def get_kristin_cut():
    try:
        row = get_db().execute('SELECT value FROM Settings WHERE key = "kristin_cut"').fetchone()
        return float(row[0]) if row else 0.10
    except Exception:
        return 0.10

# Supported display formats: (strftime_pattern, label, example)
DATE_FORMAT_OPTIONS = [
    ('%m/%d/%Y',  'MM/DD/YYYY',   'US standard — 05/09/2026'),
    ('%d/%m/%Y',  'DD/MM/YYYY',   'European — 09/05/2026'),
    ('%Y-%m-%d',  'YYYY-MM-DD',   'ISO 8601 — 2026-05-09'),
    ('%d-%m-%Y',  'DD-MM-YYYY',   'European dashes — 09-05-2026'),
    ('%d.%m.%Y',  'DD.MM.YYYY',   'European dots — 09.05.2026'),
    ('%b %d, %Y', 'Mon DD, YYYY', 'Abbreviated — May 09, 2026'),
    ('%d %b %Y',  'DD Mon YYYY',  'International — 09 May 2026'),
    ('%B %d, %Y', 'Month DD, YYYY','Long — May 09, 2026'),
]

def get_date_format():
    """Return the configured strftime date format, cached on flask g per request."""
    if hasattr(g, '_date_fmt'):
        return g._date_fmt
    try:
        row = get_db().execute('SELECT value FROM Settings WHERE key="date_format"').fetchone()
        fmt = row[0] if row else '%m/%d/%Y'
        # Validate it's one of our known patterns
        known = {o[0] for o in DATE_FORMAT_OPTIONS}
        if fmt not in known:
            fmt = '%m/%d/%Y'
    except Exception:
        fmt = '%m/%d/%Y'
    g._date_fmt = fmt
    return fmt

def split_for_account(account, country, account_splits, default_split):
    if not account:
        return default_split
    acct_lower = account.lower()
    country_upper = (country or '').strip().upper()
    for name, countries, rate in account_splits:
        if name == acct_lower:
            if not countries:
                return rate
            if country_upper in countries:
                return rate
    return default_split

# Booking types suppressed from all Team (non-Kristin) views
_TEAM_SUPPRESS_SQL = (
    "LOWER(COALESCE({alias}BookingAssociate,'')) = 'kristin house' OR "
    "LOWER(COALESCE({alias}BookingType,'')) NOT IN "
    "('other services', 'other', 'conference management', 'cm')"
)

def effective_split(associate, account, country, account_splits, default_split, kristin_split, kristin_cut):
    """Net split for a booking row, accounting for Kristin House structure."""
    if associate and associate.strip().lower() == 'kristin house':
        return kristin_split
    base = split_for_account(account, country, account_splits, default_split)
    return max(0.0, base - kristin_cut)

def split_label(associate, split, default_split, kristin_split, kristin_cut):
    """Human-readable label for the split column."""
    if associate and associate.strip().lower() == 'kristin house':
        return f"Kristin {split*100:.0f}%"
    base = split + kristin_cut   # reconstruct gross before Kristin cut
    return f"{base*100:.0f}% − {kristin_cut*100:.0f}% = {split*100:.0f}%"

# ── Dashboard ─────────────────────────────────────────────────────────────────

@app.route('/dashboard')
def dashboard():
    user = get_current_user()
    if not has_permission(user, 'dashboard'):
        flash('You do not have access to the Dashboard.', 'error')
        return redirect(url_for('index'))
    from datetime import timedelta
    db      = get_db()
    today   = datetime.today().strftime('%Y-%m-%d')
    next30  = (datetime.today() + timedelta(days=30)).strftime('%Y-%m-%d')
    year    = datetime.today().strftime('%Y')

    who = request.args.get('who', 'kristin')  # 'kristin' or 'team'
    is_kristin_view = (who == 'kristin')
    title_suffix = 'Kristin House' if is_kristin_view else 'Team Associates'

    default_split  = get_commission_split()
    account_splits = get_account_splits()
    kristin_split  = get_kristin_split()
    kristin_cut    = get_kristin_cut()

    # Bookings that received a payment this year
    paid_rows = db.execute('''
        SELECT r.BookingId AS bid, r.AccountName, r.Country,
               substr(r.StartDate,1,10) AS start_date,
               substr(r.EndDate,1,10)   AS end_date,
               COALESCE(r.USDRevenue, r.Revenue, 0) AS rev,
               r.CommissionPercent AS comm_pct,
               r.RoomRate AS room_rate,
               r.BookingAssociate AS associate,
               SUM(c.FinalPayment) AS total_paid
        FROM ReportPipeline r
        JOIN ChkRegNote c ON c.BookingID = r.BookingId
        WHERE (r.BookingStatus IS NULL OR r.BookingStatus NOT LIKE "%Cancel%")
          AND (c.Cancelled IS NULL OR c.Cancelled = 0)
          AND c.FinalPayment > 0
          AND substr(c.DateOnCheck,1,4) = ?
          AND (LOWER(COALESCE(r.BookingAssociate,'')) = 'kristin house'
               OR LOWER(COALESCE(r.BookingType,'')) NOT IN ('other services', 'other', 'conference management', 'cm'))
        GROUP BY r.BookingId
    ''', (year,)).fetchall()
    paid_bids = {r['bid'] for r in paid_rows}

    # Bookings with this-year start dates and no payment yet
    unpaid_rows = db.execute('''
        SELECT r.BookingId AS bid, r.AccountName, r.Country,
               substr(r.StartDate,1,10) AS start_date,
               substr(r.EndDate,1,10)   AS end_date,
               COALESCE(r.USDRevenue, r.Revenue, 0) AS rev,
               r.CommissionPercent AS comm_pct,
               r.RoomRate AS room_rate,
               r.BookingAssociate AS associate
        FROM ReportPipeline r
        WHERE (r.BookingStatus IS NULL OR r.BookingStatus NOT LIKE "%Cancel%")
          AND substr(r.StartDate,1,4) = ?
          AND (LOWER(COALESCE(r.BookingAssociate,'')) = 'kristin house'
               OR LOWER(COALESCE(r.BookingType,'')) NOT IN ('other services', 'other', 'conference management', 'cm'))
    ''', (year,)).fetchall()

    pickup_map = {r[0]: float(r[1] or 0) for r in db.execute('''
        SELECT BookingID, SUM(ActualPickup) FROM Pickup
        WHERE ActualPickup IS NOT NULL AND ActualPickup > 0
        GROUP BY BookingID
    ''').fetchall()}

    def _is_kristin(associate):
        return (associate or '').strip().lower() == 'kristin house'

    cats = {
        1: {'label': 'Completed & Paid',               'color': '#198754', 'count': 0, 'revenue': 0, 'value': 0},
        2: {'label': 'Future & Paid (Incentive)',       'color': '#20c997', 'count': 0, 'revenue': 0, 'value': 0},
        3: {'label': 'Completed – Has Pickup, Unpaid', 'color': '#fd7e14', 'count': 0, 'revenue': 0, 'value': 0},
        4: {'label': 'Completed – No Pickup Yet',      'color': '#dc3545', 'count': 0, 'revenue': 0, 'value': 0},
        5: {'label': 'Future / Not Paid (2026-04-10 – 2026-09-30)', 'color': '#1a3a5c', 'count': 0, 'revenue': 0, 'value': 0},
    }
    meetings_next30 = 0

    for b in paid_rows:
        if is_kristin_view != _is_kristin(b['associate']):
            continue
        end_date   = b['end_date'] or ''
        start_date = b['start_date'] or ''
        rev        = float(b['rev'] or 0)
        total_paid_amt = float(b['total_paid'] or 0)
        if start_date and today <= start_date <= next30:
            meetings_next30 += 1
        if end_date and end_date < today:
            cats[1]['count']   += 1
            cats[1]['revenue'] += rev
            cats[1]['value']   += total_paid_amt
        else:
            cats[2]['count']   += 1
            cats[2]['revenue'] += rev
            cats[2]['value']   += total_paid_amt

    for b in unpaid_rows:
        if is_kristin_view != _is_kristin(b['associate']):
            continue
        bid        = b['bid']
        is_prepaid = bid in paid_bids
        end_date   = b['end_date'] or ''
        start_date = b['start_date'] or ''
        rev        = float(b['rev'] or 0)
        comm_pct   = float(b['comm_pct'] or 0)
        room_rate  = float(b['room_rate'] or 0)
        split      = effective_split(b['associate'], b['AccountName'], b['Country'],
                                     account_splits, default_split, kristin_split, kristin_cut)
        est_comm   = rev * comm_pct * split
        if start_date and today <= start_date <= next30:
            meetings_next30 += 1
        is_done = end_date and end_date < today
        if is_prepaid and is_done:
            continue  # completed and paid — already in cat 1
        if is_kristin_view:
            if not is_prepaid and is_done and bid in pickup_map:
                actual_comm = pickup_map[bid] * room_rate * comm_pct * split if room_rate else est_comm
                cats[3]['count']   += 1; cats[3]['revenue'] += rev; cats[3]['value'] += actual_comm
            elif not is_prepaid and is_done:
                cats[4]['count']   += 1; cats[4]['revenue'] += rev; cats[4]['value'] += est_comm
            elif not start_date or start_date < '2026-10-01':
                cats[5]['count']   += 1; cats[5]['revenue'] += rev; cats[5]['value'] += est_comm
        else:
            # Team view: future only — show Kristin's 10% cut of team bookings
            if not is_done and (not start_date or start_date < '2026-10-01'):
                kristin_est = rev * comm_pct * kristin_cut
                cats[5]['count']   += 1; cats[5]['revenue'] += rev; cats[5]['value'] += kristin_est

    total_paid              = cats[1]['value'] + cats[2]['value']
    total_commission        = sum(cats[i]['value'] for i in cats)
    missing_commission_count = cats[3]['count'] + cats[4]['count']

    assoc_filter = '' if not is_kristin_view else ''
    upcoming = db.execute('''
        SELECT BookingId, EventName, Customer, AccountName, StartDate, BookingStatus,
               BookingAssociate AS associate
        FROM ReportPipeline
        WHERE substr(StartDate,1,10) >= ?
          AND (BookingStatus IS NULL OR BookingStatus NOT LIKE "%Cancel%")
          AND (LOWER(COALESCE(BookingAssociate,'')) = 'kristin house'
               OR LOWER(COALESCE(BookingType,'')) NOT IN ('other services', 'other', 'conference management', 'cm'))
        ORDER BY StartDate ASC
    ''', (today,)).fetchall()
    upcoming = [u for u in upcoming if is_kristin_view == _is_kristin(u['associate'])][:10]

    return render_template('dashboard.html',
        cats=cats, total_paid=total_paid, total_commission=total_commission,
        meetings_next30=meetings_next30, missing_commission_count=missing_commission_count,
        upcoming=upcoming, now=datetime.today(), who=who, title_suffix=title_suffix)

# ── Home / Bookings List ──────────────────────────────────────────────────────

@app.route('/')
def index():
    user = get_current_user()
    if not has_permission(user, 'bookings_view'):
        return render_template('no_access.html', section='Bookings')
    db = get_db()
    search    = request.args.get('search', '').strip()
    status    = request.args.get('status', '').strip()
    associate = request.args.get('associate', '').strip()

    query = 'SELECT * FROM ReportPipeline WHERE (BookingStatus IS NULL OR BookingStatus NOT LIKE "%Cancel%")'
    params = []

    # Account-level filter for non-admins
    accounts = get_user_account_filter(user)
    if accounts is not None:
        if accounts:
            ph = ','.join('?' * len(accounts))
            query += f' AND AccountName IN ({ph})'
            params.extend(accounts)
        else:
            query += ' AND 1=0'

    if search:
        query += ' AND (EventName LIKE ? OR "Booking Name" LIKE ? OR AccountName LIKE ? OR Customer LIKE ? OR CAST(BookingId AS TEXT) LIKE ?)'
        params += [f'%{search}%', f'%{search}%', f'%{search}%', f'%{search}%', f'%{search}%']
    if status:
        query += ' AND BookingStatus = ?'
        params.append(status)
    if associate:
        query += ' AND BookingAssociate = ?'
        params.append(associate)

    query += ' ORDER BY BookedDate DESC'
    try:
        bookings  = db.execute(query, params).fetchall()
        statuses  = [r[0] for r in db.execute('SELECT DISTINCT BookingStatus FROM ReportPipeline WHERE BookingStatus IS NOT NULL ORDER BY 1').fetchall()]
        associates = [r[0] for r in db.execute('SELECT DISTINCT BookingAssociate FROM ReportPipeline WHERE BookingAssociate IS NOT NULL ORDER BY 1').fetchall()]
    except Exception:
        bookings, statuses, associates = [], [], []

    # Calculate Kristin's commission for each booking
    kristin_split = get_kristin_split()
    kristin_cut   = get_kristin_cut()
    enriched = []
    for b in bookings:
        row = dict(b)
        total_comm = float(row.get('USDCommissionableAmount') or 0)
        assoc = (row.get('BookingAssociate') or '').strip().lower()
        if assoc == 'kristin house':
            row['kristin_comm'] = total_comm * kristin_split
        else:
            row['kristin_comm'] = total_comm * kristin_cut
        enriched.append(row)

    return render_template('index.html', bookings=enriched, search=search, status=status,
                           associate=associate, statuses=statuses, associates=associates)

# ── Booking Detail ────────────────────────────────────────────────────────────

@app.route('/booking/<int:booking_id>')
def booking_detail(booking_id):
    user = get_current_user()
    db = get_db()
    booking = db.execute('SELECT * FROM ReportPipeline WHERE BookingId = ?', (booking_id,)).fetchone()
    if not booking:
        flash('Booking not found.', 'error')
        return redirect(url_for('index'))
    acct_filter = get_user_account_filter(user)
    if acct_filter is not None and booking['AccountName'] not in acct_filter:
        flash('You do not have access to this booking.', 'error')
        return redirect(url_for('index'))
    pickups = db.execute('SELECT *, rowid FROM Pickup WHERE BookingID = ? ORDER BY rowid', (booking_id,)).fetchall()
    checks  = db.execute('SELECT * FROM ChkRegNote WHERE BookingID = ? ORDER BY DateOnCheck DESC', (booking_id,)).fetchall()
    booking_contracts = db.execute(
        'SELECT id, filename, upload_date FROM booking_contract WHERE CAST(booking_id AS TEXT)=CAST(? AS TEXT) ORDER BY upload_date DESC',
        (booking_id,)
    ).fetchall()
    default_split  = get_commission_split()
    account_splits = get_account_splits()
    kristin_split  = get_kristin_split()
    kristin_cut    = get_kristin_cut()
    split = effective_split(booking['BookingAssociate'], booking['AccountName'], booking['Country'],
                            account_splits, default_split, kristin_split, kristin_cut)
    try:
        comm_pct = float(booking['CommissionPercent'] or 0)
    except Exception:
        comm_pct = 0
    return render_template('booking_detail.html', booking=booking, pickups=pickups, checks=checks,
        comm_pct=comm_pct, split=split, booking_contracts=booking_contracts)

# ── Cancel Booking ────────────────────────────────────────────────────────────

@app.route('/booking/<int:booking_id>/cancel', methods=['POST'])
def booking_cancel(booking_id):
    db = get_db()
    db.execute('UPDATE ReportPipeline SET BookingStatus = ? WHERE BookingId = ?', ('Cancelled', booking_id))
    db.commit()
    flash('Booking has been cancelled.', 'success')
    return redirect(url_for('booking_detail', booking_id=booking_id))

# ── Booking Contracts ─────────────────────────────────────────────────────────

@app.route('/booking/<booking_id>/contract/upload', methods=['POST'])
def booking_contract_upload(booking_id):
    files = request.files.getlist('contract_files')
    db = get_db()
    last_bcid = None
    for f in files:
        if f and f.filename:
            db.execute(
                'INSERT INTO booking_contract (booking_id, filename, file_data, upload_date) VALUES (?,?,?,date("now"))',
                (booking_id, f.filename, f.read())
            )
            last_bcid = db.execute('SELECT last_insert_rowid()').fetchone()[0]
    db.commit()
    # Go straight into AI parsing flow for the last uploaded file
    if last_bcid:
        return redirect(url_for('booking_contract_extract', booking_id=booking_id, bcid=last_bcid))
    flash('No file selected.', 'error')
    return redirect(url_for('booking_detail', booking_id=booking_id))


@app.route('/booking/<booking_id>/contract/<int:bcid>/extract')
def booking_contract_extract(booking_id, bcid):
    """Route to pickup extraction — picks the right hotel or shows a selector."""
    db = get_db()
    bc = db.execute(
        'SELECT id FROM booking_contract WHERE id=? AND CAST(booking_id AS TEXT)=CAST(? AS TEXT)',
        (bcid, booking_id)
    ).fetchone()
    if not bc:
        flash('Contract not found.', 'error')
        return redirect(url_for('booking_detail', booking_id=booking_id))
    configs = db.execute(
        "SELECT * FROM pickup_config WHERE CAST(booking_id AS TEXT)=CAST(? AS TEXT) AND status != 'archived' ORDER BY id",
        (booking_id,)
    ).fetchall()
    if len(configs) == 0:
        flash('No pickup records found for this booking — add one first.', 'warning')
        return redirect(url_for('booking_detail', booking_id=booking_id))
    if len(configs) == 1:
        return redirect(url_for('booking_contract_parse_for_pickup',
                                booking_id=booking_id, bcid=bcid, cid=configs[0]['id']))
    return redirect(url_for('booking_contract_select_pickup',
                            booking_id=booking_id, bcid=bcid))


@app.route('/booking/<booking_id>/contract/<int:bcid>/select-pickup')
def booking_contract_select_pickup(booking_id, bcid):
    """Choose which pickup hotel record a booking-level contract belongs to."""
    db = get_db()
    bc = db.execute(
        'SELECT id, filename FROM booking_contract WHERE id=? AND CAST(booking_id AS TEXT)=CAST(? AS TEXT)',
        (bcid, booking_id)
    ).fetchone()
    if not bc:
        flash('Contract not found.', 'error')
        return redirect(url_for('booking_detail', booking_id=booking_id))
    configs = db.execute(
        "SELECT * FROM pickup_config WHERE CAST(booking_id AS TEXT)=CAST(? AS TEXT) AND status != 'archived' ORDER BY id",
        (booking_id,)
    ).fetchall()
    return render_template('booking_contract_select_pickup.html',
                           booking_id=booking_id, bc=bc, configs=configs)


@app.route('/booking/<booking_id>/contract/<int:bcid>/parse-for-pickup/<int:cid>')
def booking_contract_parse_for_pickup(booking_id, bcid, cid):
    """Parse a stored booking contract and show the pickup review page for a chosen hotel."""
    from pickup_utils import parse_contract_document
    import base64
    db = get_db()
    bc = db.execute(
        'SELECT filename, file_data FROM booking_contract WHERE id=? AND CAST(booking_id AS TEXT)=CAST(? AS TEXT)',
        (bcid, booking_id)
    ).fetchone()
    if not bc or not bc['file_data']:
        flash('Contract file not found.', 'error')
        return redirect(url_for('booking_detail', booking_id=booking_id))
    config = db.execute('SELECT * FROM pickup_config WHERE id=?', (cid,)).fetchone()
    if not config:
        flash('Pickup record not found.', 'error')
        return redirect(url_for('booking_detail', booking_id=booking_id))
    extracted = parse_contract_document(bc['file_data'], filename=bc['filename'])
    if extracted.get('error'):
        flash(f'Could not parse contract: {extracted["error"]}', 'error')
        return redirect(url_for('booking_detail', booking_id=booking_id))
    # file_b64 left empty — file already stored in booking_contract, no need to round-trip it
    return render_template('pickup_contract_review.html',
                           config=config, extracted=extracted,
                           filename=bc['filename'], file_b64='')

@app.route('/booking/<booking_id>/contract/<int:cid>/download')
def booking_contract_download(booking_id, cid):
    db = get_db()
    row = db.execute(
        'SELECT filename, file_data FROM booking_contract WHERE id=? AND CAST(booking_id AS TEXT)=CAST(? AS TEXT)',
        (cid, booking_id)
    ).fetchone()
    if not row:
        flash('Contract file not found.', 'error')
        return redirect(url_for('booking_detail', booking_id=booking_id))
    from flask import send_file
    import io
    return send_file(io.BytesIO(row['file_data']), download_name=row['filename'], as_attachment=True)

@app.route('/booking/<booking_id>/contract/<int:cid>/delete', methods=['POST'])
def booking_contract_delete(booking_id, cid):
    db = get_db()
    db.execute('DELETE FROM booking_contract WHERE id=? AND CAST(booking_id AS TEXT)=CAST(? AS TEXT)', (cid, booking_id))
    db.commit()
    flash('Contract deleted.', 'success')
    return redirect(url_for('booking_detail', booking_id=booking_id))

# ── New Booking ───────────────────────────────────────────────────────────────

@app.route('/booking/new', methods=['GET', 'POST'])
def booking_new():
    user = get_current_user()
    if not has_permission(user, 'bookings_edit'):
        flash('You do not have permission to add bookings.', 'error')
        return redirect(url_for('index'))
    if request.method == 'POST':
        db = get_db()
        fields = ['BookingId','BookingName','BookingAssociate','ShareType','BookingType','BookingStatus',
                  'Addendum','BookedDate','StartDate','EndDate','AccountName','EventName',
                  'Customer','Address','City','State','Country','Brand','Chain','Currency',
                  'ExchangeRate','Advance','ContractedAmount','PeakRooms','TotalRoomNights',
                  'RoomRate','Revenue','OtherRevenue','USDRate','USDRevenue',
                  'CommissionPercent','USDCommissionableAmount']
        values = [request.form.get(f, '').strip() or None for f in fields]
        placeholders = ', '.join(fields)
        qmarks = ', '.join(['?' for _ in fields])
        try:
            db.execute(f'INSERT INTO ReportPipeline ({placeholders}) VALUES ({qmarks})', values)
            db.commit()
            flash('Booking added successfully.', 'success')
            return redirect(url_for('booking_detail', booking_id=request.form.get('BookingId')))
        except Exception as e:
            flash(f'Error saving booking: {e}', 'error')
    return render_template('booking_new.html', prefill=request.args)

# ── Edit Booking ──────────────────────────────────────────────────────────────

@app.route('/booking/<int:booking_id>/edit', methods=['GET', 'POST'])
def booking_edit(booking_id):
    user = get_current_user()
    if not has_permission(user, 'bookings_edit'):
        flash('You do not have permission to edit bookings.', 'error')
        return redirect(url_for('booking_detail', booking_id=booking_id))
    db = get_db()
    booking = db.execute('SELECT * FROM ReportPipeline WHERE BookingId = ?', (booking_id,)).fetchone()
    if not booking:
        flash('Booking not found.', 'error')
        return redirect(url_for('index'))
    if request.method == 'POST':
        fields = ['BookingName','BookingAssociate','ShareType','BookingType','BookingStatus','Addendum',
                  'BookedDate','StartDate','EndDate','AccountName','EventName','Customer',
                  'Address','City','State','Country','Brand','Chain','Currency','ExchangeRate',
                  'Advance','ContractedAmount','PeakRooms','TotalRoomNights','RoomRate',
                  'Revenue','OtherRevenue','USDRate','USDRevenue','CommissionPercent',
                  'USDCommissionableAmount']
        set_clause = ', '.join([f'{f} = ?' for f in fields])
        values = [request.form.get(f, '').strip() or None for f in fields]
        values.append(booking_id)
        try:
            db.execute(f'UPDATE ReportPipeline SET {set_clause} WHERE BookingId = ?', values)
            db.commit()
            flash('Booking updated.', 'success')
            return redirect(url_for('booking_detail', booking_id=booking_id))
        except Exception as e:
            flash(f'Error updating booking: {e}', 'error')
    return render_template('booking_edit.html', booking=booking)

# ── Import Bookings ───────────────────────────────────────────────────────────

@app.route('/import', methods=['GET', 'POST'])
def import_bookings():
    user = get_current_user()
    if not has_permission(user, 'import_bookings'):
        flash('You do not have access to this import tool.', 'error')
        return redirect(url_for('index'))
    if request.method == 'POST':
        file = request.files.get('file')
        if not file or not file.filename:
            flash('No file selected.', 'error')
            return redirect(url_for('import_bookings'))
        ext = file.filename.rsplit('.', 1)[-1].lower()
        tmp_path = f'/tmp/cpainc_import.{ext}'
        file.save(tmp_path)
        try:
            if ext == 'csv':
                df = pd.read_csv(tmp_path, dtype=str, encoding='utf-8-sig')
                # Normalise column names immediately — no header-row scan needed
                df.columns = [str(c).strip() for c in df.columns]
            elif ext == 'xls':
                # ConferenceDirect exports .xls files that are sometimes HTML tables.
                # Try xlrd first; fall back to HTML parsing when xlrd rejects it.
                try:
                    df = pd.read_excel(tmp_path, engine='xlrd', header=None)
                except Exception:
                    tables = pd.read_html(tmp_path)
                    if not tables:
                        raise ValueError('No table found in file.')
                    df = tables[0].astype(str).replace('nan', '')
                    # read_html already parses <th> as column headers — check if
                    # we still need to find the header row or if columns are correct.
                if not any(str(v).strip().lower() == 'booking id' for v in df.columns):
                    header_row = None
                    for i, row in df.iterrows():
                        if any(str(v).strip().lower() == 'booking id' for v in row.values):
                            header_row = i
                            break
                    if header_row is None:
                        flash('Could not find header row with "Booking Id" in the file.', 'error')
                        return redirect(url_for('import_bookings'))
                    df.columns = df.iloc[header_row]
                    df = df.iloc[header_row + 1:].reset_index(drop=True)
            else:
                df = pd.read_excel(tmp_path, engine='openpyxl', header=None)
                header_row = None
                for i, row in df.iterrows():
                    if any(str(v).strip().lower() == 'booking id' for v in row.values):
                        header_row = i
                        break
                if header_row is None:
                    flash('Could not find header row with "Booking Id" in the file.', 'error')
                    return redirect(url_for('import_bookings'))
                df.columns = df.iloc[header_row]
                df = df.iloc[header_row + 1:].reset_index(drop=True)
            # ── Normalise column names from CSV/Excel exports ───────────────
            col_map = {
                'booking id':                               'Booking Id',
                # Event ID (parent event)
                'event: event id':                          'Event ID',
                # Associate
                'booking associate: full name':             'Booking Associate',
                'booking associate':                        'Booking Associate',
                # Event name
                'event: event name':                        'Event Name',
                'event: event  name':                       'Event Name',
                'booking name':                             'Booking Name',
                # Account / Customer
                'account name: account name':               'Account Name',
                'account name':                             'Account Name',
                'customer account name':                    'Account Name',
                'entity to invoice: account name':          'Customer',
                'entity to invoice':                        'Customer',
                # Address
                'entity to invoice: billing street':        'Address',
                'entity to invoice: billing city':          'City',
                'entity to invoice: billing state/province': 'State',
                'entity to invoice: billing country':       'Country',
                # Status & type
                'status':                                   'Booking Status',
                'type':                                     'Booking Type',
                'subtype':                                  'Booking Type',
                'addendum submitted':                       'Addendum',
                # Financial
                'total room revenue':                       'Revenue',
                'total room nights':                        'Total Room Nights',
                'peak rooms':                               'Peak Rooms',
                'contract commission %':                    'Commission Percent',
                'contract rate':                            'Room Rate',
                'gross booking value (calc) (converted)':   'Contracted Amount',
                'gross booking value':                      'Contracted Amount',
                'net contract value (converted)':           'USD Commissionable Amount',
                'total commission amount':                  'USD Commissionable Amount',
            }
            df.columns = [col_map.get(str(c).strip().lower(), str(c).strip()) for c in df.columns]
            # Deduplicate: keep first occurrence if two columns map to same name
            seen = {}
            for i, c in enumerate(df.columns):
                if c not in seen:
                    seen[c] = i
            df = df.iloc[:, sorted(seen.values())]

            bid_col = next((c for c in df.columns if c.lower() == 'booking id' or c == 'Booking Id'), None)
            if bid_col and bid_col != 'Booking Id':
                df = df.rename(columns={bid_col: 'Booking Id'})
            if 'Booking Id' not in df.columns:
                flash('Could not find Booking ID column in the file.', 'error')
                return redirect(url_for('import_bookings'))

            df = df.dropna(subset=['Booking Id'])
            df['Booking Id'] = df['Booking Id'].astype(str).str.strip()
            df = df[df['Booking Id'] != '']

            db = get_db()
            added, skipped, updated = 0, 0, 0
            currency_fields = {'Contracted Amount', 'Room Rate', 'Revenue', 'Other Revenue', 'USD Commissionable Amount'}

            fields = [
                'Booking Id', 'Event ID', 'Booking Name', 'Booking Associate', 'Share Type', 'Booking Type', 'Booking Status',
                'Addendum', 'Booked Date', 'Start Date', 'End Date', 'Account Name', 'Event Name',
                'Customer', 'Address', 'City', 'State', 'Country', 'Brand', 'Chain', 'Currency',
                'Exchange Rate', 'Advance', 'Contracted Amount', 'Peak Rooms', 'Total Room Nights',
                'Room Rate', 'Revenue', 'Other Revenue', 'USD Rate', 'USD Revenue',
                'Commission Percent', 'USD Commissionable Amount'
            ]

            for _, row in df.iterrows():
                bid = str(row.get('Booking Id', '')).strip().split('.')[0]
                if not bid:
                    continue

                values = []
                for f in fields:
                    v = row.get(f)
                    is_na = pd.isna(v) if not isinstance(v, str) else False
                    if is_na or v is None:
                        values.append(None)
                        continue
                    val = str(v).strip()
                    # Treat '-' as blank (CD uses '-' for empty fields like Event ID)
                    if val == '-':
                        values.append(None)
                        continue
                    if f in currency_fields:
                        val = val.replace('USD', '').replace('$', '').replace(',', '').strip()
                    # Normalise Commission Percent: "10.00%" → "0.10"
                    if f == 'Commission Percent' and val.endswith('%'):
                        try:
                            val = str(float(val.rstrip('%')) / 100)
                        except ValueError:
                            val = None
                    values.append(val if val else None)

                # Map from display field names to DB column names
                db_map = {
                    'Booking Id': 'BookingId', 'Event ID': 'EventID',
                    'Booking Name': 'BookingName',
                    'Booking Associate': 'BookingAssociate', 'Share Type': 'ShareType',
                    'Booking Type': 'BookingType', 'Booking Status': 'BookingStatus',
                    'Addendum': 'Addendum', 'Booked Date': 'BookedDate',
                    'Start Date': 'StartDate', 'End Date': 'EndDate',
                    'Account Name': 'AccountName', 'Event Name': 'EventName',
                    'Customer': 'Customer', 'Address': 'Address', 'City': 'City',
                    'State': 'State', 'Country': 'Country', 'Brand': 'Brand',
                    'Chain': 'Chain', 'Currency': 'Currency',
                    'Exchange Rate': 'ExchangeRate', 'Advance': 'Advance',
                    'Contracted Amount': 'ContractedAmount', 'Peak Rooms': 'PeakRooms',
                    'Total Room Nights': 'TotalRoomNights', 'Room Rate': 'RoomRate',
                    'Revenue': 'Revenue', 'Other Revenue': 'OtherRevenue',
                    'USD Rate': 'USDRate', 'USD Revenue': 'USDRevenue',
                    'Commission Percent': 'CommissionPercent',
                    'USD Commissionable Amount': 'USDCommissionableAmount',
                }
                db_fields = [db_map[f] for f in fields]

                exists = db.execute('SELECT 1 FROM ReportPipeline WHERE BookingId = ?', (bid,)).fetchone()
                if exists:
                    # Update Commission Percent and Event ID if previously missing
                    comm_idx    = fields.index('Commission Percent')
                    eventid_idx = fields.index('Event ID')
                    changed = False
                    if values[comm_idx] is not None:
                        db.execute(
                            'UPDATE ReportPipeline SET CommissionPercent = ? WHERE BookingId = ? AND (CommissionPercent IS NULL OR CommissionPercent = "")',
                            (values[comm_idx], bid)
                        )
                        if db.execute('SELECT changes()').fetchone()[0]:
                            changed = True
                    if values[eventid_idx] is not None:
                        db.execute(
                            'UPDATE ReportPipeline SET EventID = ? WHERE BookingId = ? AND (EventID IS NULL OR EventID = "")',
                            (values[eventid_idx], bid)
                        )
                        if db.execute('SELECT changes()').fetchone()[0]:
                            changed = True
                    if changed:
                        updated += 1
                    skipped += 1
                    continue

                placeholders_str = ', '.join(db_fields)
                qmarks_str = ', '.join(['?' for _ in db_fields])
                db.execute(f'INSERT INTO ReportPipeline ({placeholders_str}) VALUES ({qmarks_str})', values)
                added += 1

                # Auto-create pickup tracking entry for this new booking
                val_dict = dict(zip(db_fields, values))
                _create_pickup_config_from_booking(
                    db, bid,
                    account   = val_dict.get('AccountName'),
                    event     = val_dict.get('BookingName') or val_dict.get('EventName'),
                    hotel     = val_dict.get('Customer'),
                    start_str = val_dict.get('StartDate'),
                    end_str   = val_dict.get('EndDate'),
                    peak_rooms= val_dict.get('PeakRooms'),
                    room_rate = val_dict.get('RoomRate'),
                )

            db.commit()
            msg = f'Import complete: {added} added, {skipped} already existed (skipped).'
            if updated:
                msg += f' {updated} existing record{"s" if updated != 1 else ""} updated with missing commission % or Event ID.'
            flash(msg, 'success')
        except Exception as e:
            flash(f'Import error: {e}', 'error')
        return redirect(url_for('index'))
    return render_template('import.html')

# ── Import Cancelled Meetings ─────────────────────────────────────────────────

@app.route('/import/cancelled', methods=['GET', 'POST'])
def import_cancelled():
    user = get_current_user()
    if not has_permission(user, 'import_cancelled'):
        flash('You do not have access to this import tool.', 'error')
        return redirect(url_for('index'))
    from datetime import datetime as _dt, timedelta as _td

    def excel_date(v):
        """Convert an Excel serial date, datetime object, or date string to YYYY-MM-DD."""
        if v is None:
            return None
        # pandas may return datetime/Timestamp objects directly
        if hasattr(v, 'strftime'):
            return v.strftime('%Y-%m-%d')
        try:
            f = float(v)
            if f > 0:
                return (_dt(1899, 12, 30) + _td(days=f)).strftime('%Y-%m-%d')
        except (ValueError, TypeError):
            pass
        s = str(v).strip()
        for fmt in ('%Y-%m-%d %H:%M:%S', '%Y-%m-%d', '%m/%d/%Y', '%m/%d/%y'):
            try:
                return _dt.strptime(s, fmt).strftime('%Y-%m-%d')
            except ValueError:
                pass
        return s or None

    if request.method == 'POST':
        file = request.files.get('file')
        if not file or not file.filename:
            flash('No file selected.', 'error')
            return redirect(url_for('import_cancelled'))
        ext = file.filename.rsplit('.', 1)[-1].lower()
        tmp_path = f'/tmp/cpainc_cancelled.{ext}'
        file.save(tmp_path)
        try:
            df = pd.read_excel(tmp_path, engine='xlrd' if ext == 'xls' else 'openpyxl', header=None)
            header_row = None
            for i, row in df.iterrows():
                if any(str(v).strip() == 'Booking Id' for v in row.values):
                    header_row = i
                    break
            if header_row is None:
                flash('Could not find header row with "Booking Id" in the file.', 'error')
                return redirect(url_for('import_cancelled'))
            df.columns = df.iloc[header_row]
            df = df.iloc[header_row + 1:].reset_index(drop=True)
            df = df.dropna(subset=['Booking Id'])
            df['Booking Id'] = df['Booking Id'].astype(str).str.strip()
            df = df[df['Booking Id'] != '']

            def clean(v, is_currency=False):
                if pd.isna(v) if hasattr(pd, 'isna') else v != v:
                    return None
                s = str(v).strip()
                if is_currency:
                    s = s.replace('$', '').replace(',', '').strip()
                return s or None

            db = get_db()
            updated, inserted, skipped = 0, 0, 0

            for _, row in df.iterrows():
                bid = str(row.get('Booking Id', '')).strip().split('.')[0]
                if not bid:
                    continue

                exists = db.execute(
                    'SELECT BookingStatus FROM ReportPipeline WHERE BookingId = ?', (bid,)
                ).fetchone()

                if exists:
                    current_status = (exists['BookingStatus'] or '').strip()
                    if 'cancel' in current_status.lower():
                        skipped += 1
                        continue
                    db.execute(
                        'UPDATE ReportPipeline SET BookingStatus = ? WHERE BookingId = ?',
                        ('Cancelled', bid)
                    )
                    db.execute(
                        'UPDATE ChkRegNote SET Cancelled = 1 WHERE BookingID = ? AND (Cancelled IS NULL OR Cancelled = 0)',
                        (bid,)
                    )
                    updated += 1
                else:
                    db.execute('''INSERT INTO ReportPipeline
                        (BookingId,BookingAssociate,ShareType,BookingType,BookingStatus,
                         Addendum,BookedDate,StartDate,EndDate,AccountName,EventName,
                         Customer,Address,City,State,Country,Brand,Chain,Currency,
                         ExchangeRate,Advance,ContractedAmount,PeakRooms,TotalRoomNights,
                         RoomRate,Revenue,OtherRevenue,USDRate,USDRevenue,
                         CommissionPercent,USDCommissionableAmount)
                        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''', (
                        bid,
                        clean(row.get('Booking Associate')),
                        clean(row.get('Share Type')),
                        clean(row.get('Booking Type')),
                        'Cancelled',
                        clean(row.get('Addendum')),
                        excel_date(row.get('Booked Date')),
                        excel_date(row.get('Start Date')),
                        excel_date(row.get('End Date')),
                        clean(row.get('Account Name')),
                        clean(row.get('Event Name')),
                        clean(row.get('Customer')),
                        clean(row.get('Address')),
                        clean(row.get('City')),
                        clean(row.get('State')),
                        clean(row.get('Country')),
                        clean(row.get('Brand')),
                        clean(row.get('Chain')),
                        clean(row.get('Currency')),
                        clean(row.get('Exchange Rate')),
                        clean(row.get('Advance')),
                        clean(row.get('Contracted Amount'), True),
                        clean(row.get('Peak Rooms')),
                        clean(row.get('Total Room Nights')),
                        clean(row.get('Room Rate'), True),
                        clean(row.get('Revenue'), True),
                        clean(row.get('Other Revenue'), True),
                        clean(row.get('USD Rate')),
                        clean(row.get('USD Revenue'), True),
                        clean(row.get('Commission Percent')),
                        clean(row.get('USD Commissionable Amount'), True),
                    ))
                    inserted += 1

            db.commit()
            flash(
                f'Cancelled import complete: {updated} updated to Cancelled, '
                f'{inserted} new bookings added, {skipped} already Cancelled (skipped).',
                'success'
            )
        except Exception as e:
            flash(f'Import error: {e}', 'error')
        return redirect(url_for('index'))
    return render_template('import_cancelled.html')

# ── Import Payments ───────────────────────────────────────────────────────────

@app.route('/import/payments', methods=['GET', 'POST'])
def import_payments():
    user = get_current_user()
    if not has_permission(user, 'import_payments'):
        flash('You do not have access to this import tool.', 'error')
        return redirect(url_for('index'))
    if request.method == 'POST':
        file = request.files.get('file')
        if not file or not file.filename:
            flash('No file selected.', 'error')
            return redirect(url_for('import_payments'))
        ext = file.filename.rsplit('.', 1)[-1].lower()
        tmp_path = f'/tmp/cpainc_payments.{ext}'
        file.save(tmp_path)
        try:
            if ext == 'csv':
                df = pd.read_csv(tmp_path, dtype=str, encoding='utf-8-sig')
                df.columns = [str(c).strip() for c in df.columns]
                # Normalize common CSV column name variants → expected names
                col_map = {
                    'Booking ID': 'Booking Number',
                    'BookingID': 'Booking Number',
                    'Check Number': 'Reference',
                    'Check #': 'Reference',
                    'Payment Amount': 'Amount',
                    'Payment Date': 'Date',
                }
                df.rename(columns=col_map, inplace=True)
                if 'Booking Number' not in df.columns:
                    flash('CSV must contain a "Booking Number" or "Booking ID" column.', 'error')
                    return redirect(url_for('import_payments'))
                df = df[pd.to_numeric(df['Booking Number'], errors='coerce').notna()]
                df['Booking Number'] = df['Booking Number'].astype(str).str.strip().str.split('.').str[0]
            else:
                df = pd.read_excel(tmp_path, engine='xlrd' if ext == 'xls' else 'openpyxl', header=None)
                header_row = None
                for i, row in df.iterrows():
                    if any(str(v).strip() == 'Booking Number' for v in row.values):
                        header_row = i
                        break
                if header_row is None:
                    flash('Could not find header row with "Booking Number" in the file.', 'error')
                    return redirect(url_for('import_payments'))

                df.columns = df.iloc[header_row]
                df = df.iloc[header_row + 1:].reset_index(drop=True)
                df = df[pd.to_numeric(df['Booking Number'], errors='coerce').notna()]
                df['Booking Number'] = df['Booking Number'].astype(str).str.strip().str.split('.').str[0]

            db = get_db()
            today = datetime.now().strftime('%Y-%m-%d')
            added, skipped, new_bookings = 0, 0, []

            for _, row in df.iterrows():
                booking_id = str(row.get('Booking Number', '')).strip()
                if not booking_id:
                    continue

                raw_date = row.get('Date')
                if pd.isna(raw_date) if hasattr(pd, 'isna') else raw_date != raw_date:
                    payment_date = None
                else:
                    try:
                        payment_date = pd.to_datetime(raw_date).strftime('%Y-%m-%d')
                    except Exception:
                        payment_date = str(raw_date)[:10]

                check_num = str(row.get('Reference', '') or '').strip() or None
                amount = row.get('Amount')
                if pd.isna(amount) if hasattr(pd, 'isna') else amount != amount:
                    amount = None

                if payment_date:
                    dup = db.execute(
                        'SELECT 1 FROM ChkRegNote WHERE BookingID = ? AND DateOnCheck LIKE ? AND FinalPayment = ?',
                        (booking_id, f'{payment_date}%', amount)
                    ).fetchone()
                    if dup:
                        skipped += 1
                        continue

                invoice = str(row.get('Invoice Number', '') or '').upper()
                is_advance = 1 if ('HI INC' in invoice or 'MI INC' in invoice) else 0

                booking_exists = db.execute('SELECT 1 FROM ReportPipeline WHERE BookingId = ?', (booking_id,)).fetchone()
                if not booking_exists:
                    def safe_date(val):
                        try:
                            if pd.isna(val): return None
                            return pd.to_datetime(val).strftime('%Y-%m-%d')
                        except Exception:
                            return None
                    db.execute('''INSERT INTO ReportPipeline
                        (BookingId,BookingAssociate,BookingType,AccountName,EventName,Customer,StartDate,EndDate,BookingStatus)
                        VALUES (?,?,?,?,?,?,?,?,"Definite")''', (
                        booking_id,
                        str(row.get('Booking Associate', '') or '').strip() or None,
                        str(row.get('Booking Type', '') or '').strip() or None,
                        str(row.get('Account Name', '') or '').strip() or None,
                        str(row.get('Event Name', '') or '').strip() or None,
                        str(row.get('Hotel', '') or '').strip() or None,
                        safe_date(row.get('Start Date')),
                        safe_date(row.get('End Date')),
                    ))
                    new_bookings.append({
                        'booking_id': booking_id,
                        'account':    str(row.get('Account Name', '') or '').strip(),
                        'event':      str(row.get('Event Name', '') or '').strip(),
                        'hotel':      str(row.get('Hotel', '') or '').strip(),
                        'amount':     amount,
                    })

                db.execute('''INSERT INTO ChkRegNote
                    (BookingID,FinalPayment,Check_,DateOnCheck,EntryDate,Cancelled,AuditFlag,Advance)
                    VALUES (?,?,?,?,?,0,0,?)''',
                    (booking_id, amount, check_num, payment_date, today, is_advance))
                added += 1

            db.commit()
            return render_template('import_payments_result.html',
                                   added=added, skipped=skipped, new_bookings=new_bookings)
        except Exception as e:
            flash(f'Import error: {e}', 'error')
            return redirect(url_for('import_payments'))
    return render_template('import_payments.html')

# ── Import Payment Voucher (PDF) ──────────────────────────────────────────────

@app.route('/import/voucher', methods=['GET', 'POST'])
def import_voucher():
    user = get_current_user()
    if not has_permission(user, 'import_payments'):
        flash('You do not have access to this import tool.', 'error')
        return redirect(url_for('index'))

    if request.method == 'GET':
        return render_template('import_voucher.html')

    # ── Parse uploaded PDFs ───────────────────────────────────────────────────
    files = request.files.getlist('files')
    files = [f for f in files if f and f.filename]
    if not files:
        flash('No files selected.', 'error')
        return redirect(url_for('import_voucher'))
    bad = [f.filename for f in files if not f.filename.lower().endswith('.pdf')]
    if bad:
        flash(f'Only PDF files are supported: {", ".join(bad)}', 'error')
        return redirect(url_for('import_voucher'))

    import pdfplumber, re, io as _io

    def _parse_voucher_pdf(raw_bytes):
        """Parse one Payment Voucher PDF. Returns a dict with header fields + booking_rows.

        Line-item approach: each invoice number line in the raw text ends with
        an optional end-date and two amounts, e.g.:
            237264-F1-TD 2026 Green Dot Corp ... 3/12/2026 450.66 450.66
            140699-F1 19.80 19.80
        We regex-scan every line of the extracted text for this pattern.
        """
        all_text = ''
        with pdfplumber.open(io.BytesIO(raw_bytes)) as pdf:
            for page in pdf.pages:
                all_text += (page.extract_text() or '') + '\n'

        def _find(pattern, text, group=1, flags=re.IGNORECASE):
            m = re.search(pattern, text, flags)
            return m.group(group).strip() if m else None

        def _parse_amount(s):
            if not s:
                return None
            try:
                return float(str(s).replace(',', '').strip())
            except Exception:
                return None

        raw_date    = _find(r'Date\s+([\d/]+)', all_text)
        amt_usd_str = _find(r'Amount Paid in USD\s+([\d,\.]+)', all_text)
        currency    = _find(r'Currency\s+([^\n]+)', all_text) or 'USD'
        reference   = _find(r'Reference\s+([\S]+)', all_text)

        payment_date = None
        if raw_date:
            try:
                payment_date = datetime.strptime(raw_date.strip(), '%m/%d/%Y').strftime('%Y-%m-%d')
            except Exception:
                payment_date = raw_date[:10] if raw_date else None

        amt_usd = _parse_amount(amt_usd_str)
        currency_lower = currency.strip().lower()
        is_usd = currency_lower in (
            'usd', 'us dollar', 'us dollars', 'united states dollar', 'united states dollars', ''
        )

        # ── Segment-based parser ─────────────────────────────────────────────
        # Strategy: find every invoice number position, then for each one
        # grab the text up to the NEXT invoice number.  Within that segment
        # the last two decimal numbers are the amount / amount-paid, and any
        # MM/DD/YYYY is the end date.  This handles rows whose hotel/account
        # text wraps across multiple lines.
        inv_re   = re.compile(r'\d{4,9}-F\d+(?:-[A-Z0-9]+)?')
        date_re  = re.compile(r'\b(\d{1,2}/\d{1,2}/\d{4})\b')
        amt_re   = re.compile(r'[\d,]+\.\d{2}')

        positions = [(m.start(), m.group()) for m in inv_re.finditer(all_text)]

        booking_rows = []
        seen = set()
        for idx, (pos, invoice_raw) in enumerate(positions):
            if invoice_raw in seen:
                continue
            seen.add(invoice_raw)

            # Segment runs from this invoice number to the next (or end of text)
            next_pos = positions[idx + 1][0] if idx + 1 < len(positions) else len(all_text)
            segment  = all_text[pos:next_pos]

            # Last two decimal amounts in the segment = amount, amt_paid
            amounts = amt_re.findall(segment)
            if len(amounts) < 2:
                continue
            amt_raw      = amounts[-2]
            amt_paid_raw = amounts[-1]

            # First date found in segment (if any)
            end_date = None
            dm = date_re.search(segment)
            if dm:
                try:
                    end_date = datetime.strptime(dm.group(1), '%m/%d/%Y').strftime('%Y-%m-%d')
                except Exception:
                    pass

            num_match  = re.match(r'^(\d+)', invoice_raw)
            booking_id = int(num_match.group(1)) if num_match else None

            booking_rows.append({
                'booking_id':  booking_id,
                'invoice_raw': invoice_raw,
                'bname':       '',
                'account':     '',
                'hotel':       '',
                'btype':       '',
                'end_date':    end_date,
                'amount':      _parse_amount(amt_raw),
                'amt_paid':    _parse_amount(amt_paid_raw),
            })

        return {
            'payment_date': payment_date,
            'reference':    reference,
            'currency':     currency,
            'amt_usd':      amt_usd,
            'is_usd':       is_usd,
            'booking_rows': booking_rows,
        }

    # ── Process every uploaded file ───────────────────────────────────────────
    try:
        db    = get_db()
        today = datetime.now().strftime('%Y-%m-%d')
        total_added, total_skipped = 0, 0
        all_errors, all_new_bookings, file_summaries = [], [], []

        for f in files:
            raw_bytes = f.read()
            try:
                v = _parse_voucher_pdf(raw_bytes)
            except Exception as e:
                all_errors.append(f'{f.filename}: parse error — {e}')
                file_summaries.append({'filename': f.filename, 'added': 0, 'skipped': 0, 'error': str(e)})
                continue

            if not v['booking_rows']:
                all_errors.append(f'{f.filename}: no booking table found')
                file_summaries.append({'filename': f.filename, 'added': 0, 'skipped': 0, 'error': 'No booking table found'})
                continue

            payment_date = v['payment_date']
            reference    = v['reference']
            currency     = v['currency']
            amt_usd      = v['amt_usd']
            is_usd       = v['is_usd']
            booking_rows = v['booking_rows']

            total_orig = sum(r['amt_paid'] or r['amount'] or 0 for r in booking_rows) or 1

            def row_usd(row):
                if is_usd:
                    return row['amt_paid'] or row['amount']
                row_orig = row['amt_paid'] or row['amount'] or 0
                if amt_usd and total_orig:
                    return round(amt_usd * (row_orig / total_orig), 2)
                return row['amount']

            file_added, file_skipped = 0, 0

            for row in booking_rows:
                bid = row['booking_id']
                if not bid:
                    all_errors.append(f"{f.filename}: could not parse BookingID from '{row['invoice_raw']}'")
                    continue

                usd_amount  = row_usd(row)
                invoice_num = row['invoice_raw']   # e.g. "237264-F1-TD"
                check_num   = reference            # voucher reference e.g. "00001176/106"

                # Duplicate: same BookingID + same check/reference + same invoice line,
                # OR same BookingID + same date + same amount (catch re-imports)
                dup = db.execute(
                    '''SELECT 1 FROM ChkRegNote
                       WHERE BookingID = ?
                         AND (
                           (Check_ = ? AND SpecialNotes LIKE ?)
                           OR (DateOnCheck LIKE ? AND ABS(COALESCE(FinalPayment,0) - ?) < 0.02)
                         )''',
                    (bid, check_num, f'%{invoice_num}%', f'{payment_date}%', usd_amount or 0)
                ).fetchone()
                if dup:
                    file_skipped += 1
                    continue

                notes_parts = []
                notes_parts.append(f'Invoice: {invoice_num}')
                if not is_usd:
                    notes_parts.append(f'Currency: {currency}')
                if not is_usd and row['amt_paid']:
                    notes_parts.append(f'Original: {row["amt_paid"]:,.2f}')
                notes_parts.append('Voucher Import')
                special_notes = ' | '.join(notes_parts)

                booking_exists = db.execute('SELECT 1 FROM ReportPipeline WHERE BookingId=?', (bid,)).fetchone()
                if not booking_exists:
                    db.execute('''INSERT INTO ReportPipeline
                        (BookingId, BookingType, AccountName, EventName, Customer, EndDate, BookingStatus)
                        VALUES (?,?,?,?,?,?,"Definite")''',
                        (bid, row['btype'] or None, row['account'] or None,
                         row['bname'] or None, row['hotel'] or None, row['end_date']))
                    all_new_bookings.append({
                        'booking_id': bid,
                        'account':    row['account'],
                        'event':      row['bname'],
                        'hotel':      row['hotel'],
                    })

                db.execute('''INSERT INTO ChkRegNote
                    (BookingID, FinalPayment, Check_, DateOnCheck, EntryDate, SpecialNotes, Cancelled, AuditFlag, Advance)
                    VALUES (?,?,?,?,?,?,0,0,0)''',
                    (bid, usd_amount, check_num, payment_date, today, special_notes))
                file_added += 1

            total_added   += file_added
            total_skipped += file_skipped
            file_summaries.append({
                'filename':     f.filename,
                'added':        file_added,
                'skipped':      file_skipped,
                'payment_date': payment_date,
                'reference':    reference,
                'currency':     currency,
                'amt_usd':      amt_usd,
                'error':        None,
            })

        db.commit()

        return render_template('import_voucher_result.html',
            added=total_added, skipped=total_skipped,
            errors=all_errors, new_bookings=all_new_bookings,
            file_summaries=file_summaries,
            file_count=len(files))

    except Exception as e:
        flash(f'Import error: {e}', 'error')
        return redirect(url_for('import_voucher'))


# ── Pickup New ────────────────────────────────────────────────────────────────

@app.route('/booking/<int:booking_id>/pickup/new', methods=['GET', 'POST'])
def pickup_new(booking_id):
    db = get_db()
    booking = db.execute('SELECT * FROM ReportPipeline WHERE BookingId = ?', (booking_id,)).fetchone()
    if not booking:
        flash('Booking not found.', 'error')
        return redirect(url_for('index'))
    if request.method == 'POST':
        actual_pickup   = request.form.get('actual_pickup', '').strip() or None
        attrition       = request.form.get('attrition', '').strip() or None
        cutoff          = request.form.get('cutoff', '').strip() or None
        brand           = request.form.get('brand', '').strip() or None
        total_revenue   = request.form.get('total_revenue', '').strip() or None
        try:
            avg_rate = None
            if actual_pickup and total_revenue:
                try:
                    avg_rate = float(total_revenue) / float(actual_pickup)
                except Exception:
                    pass
            db.execute(
                'INSERT INTO Pickup (BookingID, ActualPickup, AttritionPercentage, CutOffDate, Brand, TotalRevenue) VALUES (?,?,?,?,?,?)',
                (booking_id, actual_pickup, attrition, cutoff, brand, total_revenue)
            )
            if avg_rate is not None:
                db.execute('UPDATE ReportPipeline SET RoomRate = ? WHERE BookingId = ?', (avg_rate, booking_id))
            db.commit()
            flash('Pickup record added.', 'success')
            return redirect(url_for('booking_detail', booking_id=booking_id))
        except Exception as e:
            flash(f'Error saving pickup: {e}', 'error')
    default_split  = get_commission_split()
    account_splits = get_account_splits()
    kristin_split  = get_kristin_split()
    kristin_cut    = get_kristin_cut()
    split = effective_split(booking['BookingAssociate'], booking['AccountName'], booking['Country'],
                            account_splits, default_split, kristin_split, kristin_cut)
    try:
        comm_pct = float(booking['CommissionPercent'] or 0)
    except Exception:
        comm_pct = 0
    return render_template('pickup_new.html', booking=booking,
        now=datetime.now().strftime('%Y-%m-%d'), comm_pct=comm_pct, split=split)

# ── Pickup Edit ───────────────────────────────────────────────────────────────

@app.route('/pickup/<int:pickup_id>/edit', methods=['GET', 'POST'])
def pickup_edit(pickup_id):
    db = get_db()
    pickup = db.execute('SELECT *, rowid FROM Pickup WHERE rowid = ?', (pickup_id,)).fetchone()
    if not pickup:
        flash('Pickup record not found.', 'error')
        return redirect(url_for('index'))
    booking_id = pickup['BookingID']
    if request.method == 'POST':
        actual_pickup   = request.form.get('actual_pickup', '').strip() or None
        attrition       = request.form.get('attrition', '').strip() or None
        cutoff          = request.form.get('cutoff', '').strip() or None
        brand           = request.form.get('brand', '').strip() or None
        total_revenue   = request.form.get('total_revenue', '').strip() or None
        try:
            avg_rate = None
            if actual_pickup and total_revenue:
                try:
                    avg_rate = float(total_revenue) / float(actual_pickup)
                except Exception:
                    pass
            db.execute(
                'UPDATE Pickup SET ActualPickup=?, AttritionPercentage=?, CutOffDate=?, Brand=?, TotalRevenue=? WHERE rowid=?',
                (actual_pickup, attrition, cutoff, brand, total_revenue, pickup_id)
            )
            if avg_rate is not None:
                db.execute('UPDATE ReportPipeline SET RoomRate = ? WHERE BookingId = ?', (avg_rate, booking_id))
            db.commit()
            flash('Pickup updated.', 'success')
            return redirect(url_for('booking_detail', booking_id=booking_id))
        except Exception as e:
            flash(f'Error updating pickup: {e}', 'error')
    booking = db.execute('SELECT * FROM ReportPipeline WHERE BookingId = ?', (pickup['BookingID'],)).fetchone()
    default_split  = get_commission_split()
    account_splits = get_account_splits()
    kristin_split  = get_kristin_split()
    kristin_cut    = get_kristin_cut()
    split = effective_split(
        booking['BookingAssociate'] if booking else None,
        booking['AccountName'] if booking else None,
        booking['Country'] if booking else None,
        account_splits, default_split, kristin_split, kristin_cut)
    try:
        comm_pct = float((booking['CommissionPercent'] if booking else None) or 0)
    except Exception:
        comm_pct = 0
    return render_template('pickup_edit.html', pickup=pickup, comm_pct=comm_pct, split=split)

# ── Pickup Delete ─────────────────────────────────────────────────────────────

@app.route('/pickup/<int:pickup_id>/delete', methods=['POST'])
def pickup_delete(pickup_id):
    db = get_db()
    pickup = db.execute('SELECT BookingID FROM Pickup WHERE rowid = ?', (pickup_id,)).fetchone()
    if pickup:
        booking_id = pickup['BookingID']
        db.execute('DELETE FROM Pickup WHERE rowid = ?', (pickup_id,))
        db.commit()
        flash('Pickup record deleted.', 'success')
    else:
        booking_id = None
        flash('Pickup record not found.', 'error')
    if booking_id:
        return redirect(url_for('booking_detail', booking_id=booking_id))
    return redirect(url_for('index'))

# ── Check New ─────────────────────────────────────────────────────────────────

@app.route('/booking/<int:booking_id>/check/new', methods=['GET', 'POST'])
def check_new(booking_id):
    db = get_db()
    booking = db.execute('SELECT * FROM ReportPipeline WHERE BookingId = ?', (booking_id,)).fetchone()
    if not booking:
        flash('Booking not found.', 'error')
        return redirect(url_for('index'))
    if request.method == 'POST':
        try:
            db.execute('''INSERT INTO ChkRegNote
                (BookingID, FinalPayment, Check_, DateOnCheck, SpecialNotes,
                 Cancelled, AuditFlag, EntryDate, DepositDate)
                VALUES (?,?,?,?,?,?,?,?,?)''', (
                booking_id,
                request.form.get('final_payment', '').strip() or None,
                request.form.get('check_number', '').strip() or None,
                request.form.get('date_on_check', '').strip() or None,
                request.form.get('special_notes', '').strip() or None,
                1 if request.form.get('cancelled') else 0,
                1 if request.form.get('audit_flag') else 0,
                datetime.now().strftime('%Y-%m-%d'),
                request.form.get('deposit_date', '').strip() or None,
            ))
            db.commit()
            flash('Payment record added.', 'success')
            return redirect(url_for('booking_detail', booking_id=booking_id))
        except Exception as e:
            flash(f'Error saving payment: {e}', 'error')
    return render_template('check_new.html', booking=booking)

# ── Check Edit ────────────────────────────────────────────────────────────────

@app.route('/check/<int:check_id>/edit', methods=['GET', 'POST'])
def check_edit(check_id):
    db = get_db()
    check = db.execute('SELECT * FROM ChkRegNote WHERE ChkRegID = ?', (check_id,)).fetchone()
    if not check:
        flash('Payment record not found.', 'error')
        return redirect(url_for('index'))
    booking_id = check['BookingID']
    if request.method == 'POST':
        try:
            db.execute('''UPDATE ChkRegNote
                SET FinalPayment=?, Check_=?, DateOnCheck=?, SpecialNotes=?,
                    Cancelled=?, AuditFlag=?, DepositDate=?
                WHERE ChkRegID=?''', (
                request.form.get('final_payment', '').strip() or None,
                request.form.get('check_number', '').strip() or None,
                request.form.get('date_on_check', '').strip() or None,
                request.form.get('special_notes', '').strip() or None,
                1 if request.form.get('cancelled') else 0,
                1 if request.form.get('audit_flag') else 0,
                request.form.get('deposit_date', '').strip() or None,
                check_id,
            ))
            db.commit()
            flash('Payment updated.', 'success')
            return redirect(url_for('booking_detail', booking_id=booking_id))
        except Exception as e:
            flash(f'Error updating payment: {e}', 'error')
    return render_template('check_edit.html', check=check)

# ── Commission Report ─────────────────────────────────────────────────────────

@app.route('/reports/commission')
def report_commission():
    user = get_current_user()
    who_param = request.args.get('who', 'team')
    if who_param == 'kristin' and not has_permission(user, 'reports_commission_kristin'):
        flash('You do not have access to that report.', 'error')
        return redirect(url_for('index'))
    if who_param == 'team' and not has_permission(user, 'reports_commission_team'):
        flash('You do not have access to that report.', 'error')
        return redirect(url_for('index'))
    date_from = request.args.get('date_from', '').strip()
    date_to   = request.args.get('date_to', '').strip()
    who       = who_param
    today     = datetime.now().strftime('%Y-%m-%d')
    rows, totals = [], {}

    if date_from and date_to:
        db = get_db()
        query = '''
            SELECT
                r.BookingId          AS booking_id,
                r.AccountName        AS account,
                r.Customer           AS hotel,
                r.BookingAssociate   AS associate,
                r.BookedDate         AS booked_date,
                r.StartDate          AS start_date,
                r.EndDate            AS end_date,
                r.RoomRate           AS room_rate,
                r.CommissionPercent  AS comm_pct,
                r.Revenue            AS revenue,
                r.USDRevenue         AS usd_revenue,
                r.TotalRoomNights    AS total_room_nights,
                r.Country            AS country,
                COALESCE(p.total_pickup, 0) AS actual_pickup,
                CASE WHEN pay.pay_count > 0 THEN 1 ELSE 0 END AS is_paid,
                COALESCE(pay.total_paid, 0) AS prepaid
            FROM ReportPipeline r
            LEFT JOIN (
                SELECT BookingID, SUM(ActualPickup) AS total_pickup
                FROM Pickup GROUP BY BookingID
            ) p ON p.BookingID = r.BookingId
            LEFT JOIN (
                SELECT BookingID, COUNT(*) AS pay_count,
                       SUM(COALESCE(FinalPayment, 0)) AS total_paid
                FROM ChkRegNote
                WHERE (Cancelled IS NULL OR Cancelled = 0)
                GROUP BY BookingID
            ) pay ON pay.BookingID = r.BookingId
            WHERE DATE(r.StartDate) BETWEEN ? AND ?
            AND (pay.pay_count IS NULL OR pay.pay_count = 0 OR DATE(r.StartDate) > ?)
            AND (r.BookingStatus IS NULL OR r.BookingStatus != 'Cancelled')
            AND NOT EXISTS (
                SELECT 1 FROM ChkRegNote c2
                WHERE c2.BookingID = r.BookingId AND c2.Cancelled = 1
            )
            AND (LOWER(COALESCE(r.BookingAssociate,'')) = 'kristin house'
                 OR LOWER(COALESCE(r.BookingType,'')) NOT IN ('other services', 'other', 'conference management', 'cm'))
            ORDER BY r.StartDate
        '''
        rows = [dict(r) for r in db.execute(query, (date_from, date_to, today)).fetchall()]

        default_split  = get_commission_split()
        account_splits = get_account_splits()
        kristin_split  = get_kristin_split()
        kristin_cut    = get_kristin_cut()
        for r in rows:
            split = effective_split(r.get('associate'), r.get('account'), r.get('country'),
                                    account_splits, default_split, kristin_split, kristin_cut)
            r['split_pct'] = split_label(r.get('associate'), split, default_split, kristin_split, kristin_cut)
            try:
                rev = float(r['usd_revenue'] or 0) or float(r['revenue'] or 0)
                r['est_commission'] = rev * float(r['comm_pct'] or 0) * split
            except Exception:
                r['est_commission'] = 0
            try:
                r['actual_commission'] = float(r['actual_pickup'] or 0) * float(r['room_rate'] or 0) * float(r['comm_pct'] or 0) * split
            except Exception:
                r['actual_commission'] = 0

        if who == 'kristin':
            rows = [r for r in rows if (r.get('associate') or '').strip().lower() == 'kristin house']
        else:
            rows = [r for r in rows if (r.get('associate') or '').strip().lower() != 'kristin house']

        totals = {
            'est_commission':    sum(float(r['est_commission'] or 0) for r in rows),
            'actual_commission': sum(r['actual_commission'] for r in rows),
            'prepaid':           sum(float(r.get('prepaid') or 0) for r in rows),
        }

    title = 'Kristin House — Missing Commission' if who == 'kristin' else 'Team — Missing Commission'
    return render_template('report_commission.html',
                           rows=rows, totals=totals, title=title, who=who,
                           date_from=date_from, date_to=date_to, today=today)

# ── Payment Report ────────────────────────────────────────────────────────────

@app.route('/reports/payments')
def report_payments():
    user = get_current_user()
    if not has_permission(user, 'reports_payments'):
        flash('You do not have access to the Payment Report.', 'error')
        return redirect(url_for('index'))
    date_from = request.args.get('date_from', '').strip()
    date_to   = request.args.get('date_to', '').strip()
    today     = datetime.now().strftime('%Y-%m-%d')
    rows, totals = [], {}

    if date_from and date_to:
        db = get_db()
        query = '''
            SELECT
                r.BookingId          AS booking_id,
                r.AccountName        AS account,
                r.EventName          AS event_name,
                r.Customer           AS hotel,
                r.StartDate          AS start_date,
                r.RoomRate           AS room_rate,
                r.CommissionPercent  AS comm_pct,
                r.Revenue            AS revenue,
                r.USDRevenue         AS usd_revenue,
                r.Country            AS country,
                r.BookingAssociate   AS associate,
                c.DateOnCheck        AS date_on_check,
                c.Check_             AS check_number,
                c.FinalPayment       AS final_payment,
                c.ChkRegID           AS chk_id,
                c.Advance            AS is_advance,
                COALESCE(p.total_pickup, 0) AS actual_pickup,
                COALESCE(all_pay.pay_count, 0) AS total_pay_count,
                COALESCE(all_pay.pay_total, 0) AS total_pay_amount
            FROM ChkRegNote c
            JOIN ReportPipeline r ON r.BookingId = c.BookingID
            LEFT JOIN (
                SELECT BookingID, SUM(ActualPickup) AS total_pickup
                FROM Pickup GROUP BY BookingID
            ) p ON p.BookingID = r.BookingId
            LEFT JOIN (
                SELECT BookingID, COUNT(*) AS pay_count, SUM(COALESCE(FinalPayment,0)) AS pay_total
                FROM ChkRegNote WHERE (Cancelled IS NULL OR Cancelled = 0)
                GROUP BY BookingID
            ) all_pay ON all_pay.BookingID = c.BookingID
            WHERE DATE(c.DateOnCheck) BETWEEN ? AND ?
            AND (c.Cancelled IS NULL OR c.Cancelled = 0)
            AND (LOWER(COALESCE(r.BookingAssociate,'')) = 'kristin house'
                 OR LOWER(COALESCE(r.BookingType,'')) NOT IN ('other services', 'other', 'conference management', 'cm'))
            ORDER BY c.DateOnCheck DESC
        '''
        raw = db.execute(query, (date_from, date_to)).fetchall()

        default_split  = get_commission_split()
        account_splits = get_account_splits()
        tolerance      = get_payment_tolerance()
        kristin_split  = get_kristin_split()
        kristin_cut    = get_kristin_cut()

        for r in raw:
            r = dict(r)
            is_kristin = (r.get('associate') or '').strip().lower() == 'kristin house'
            r['is_kristin'] = is_kristin
            final = float(r['final_payment'] or 0)
            r['final_payment_f'] = final
            r['is_future'] = r.get('start_date', '') and str(r['start_date'])[:10] > today
            total_paid  = float(r.get('total_pay_amount') or 0)
            total_count = int(r.get('total_pay_count') or 0)
            prior_total = total_paid - final
            prior_count = total_count - 1
            r['prior_count']      = prior_count if prior_count > 0 else 0
            r['prior_total']      = prior_total if prior_count > 0 else 0
            r['grand_total_paid'] = total_paid
            if is_kristin:
                split = effective_split(r.get('associate'), r.get('account'), r.get('country'),
                                        account_splits, default_split, kristin_split, kristin_cut)
                r['split_pct'] = split_label(r.get('associate'), split, default_split, kristin_split, kristin_cut)
                try:
                    actual = float(r['actual_pickup'] or 0) * float(r['room_rate'] or 0) * float(r['comm_pct'] or 0) * split
                    if actual == 0:
                        rev = float(r['usd_revenue'] or 0) or float(r['revenue'] or 0)
                        actual = rev * float(r['comm_pct'] or 0) * split
                except Exception:
                    actual = 0
                r['actual_commission'] = actual
                if actual > 0 and final < actual:
                    r['out_of_tolerance'] = (actual - final) / actual > tolerance
                else:
                    r['out_of_tolerance'] = False
            else:
                r['split_pct']        = None
                r['actual_commission'] = None
                r['out_of_tolerance'] = False
            rows.append(r)

        totals = {
            'final_payment':     sum(r['final_payment_f'] for r in rows),
            'actual_commission': sum(r['actual_commission'] or 0 for r in rows),
            'future_count':      sum(1 for r in rows if r['is_future']),
            'tolerance_count':   sum(1 for r in rows if r['out_of_tolerance']),
        }

    return render_template('report_payments.html',
                           rows=rows, totals=totals,
                           date_from=date_from, date_to=date_to, today=today,
                           tolerance=get_payment_tolerance())

# ── Payment Report Export ─────────────────────────────────────────────────────

@app.route('/reports/payments/export')
def report_payments_export():
    date_from = request.args.get('date_from', '').strip()
    date_to   = request.args.get('date_to', '').strip()
    today     = datetime.now().strftime('%Y-%m-%d')
    if not date_from or not date_to:
        flash('Please run the report with a date range before exporting.', 'error')
        return redirect(url_for('report_payments'))

    db = get_db()
    query = '''
        SELECT r.BookingId AS booking_id, r.AccountName AS account, r.EventName AS event_name,
               r.Customer AS hotel, r.StartDate AS start_date, r.RoomRate AS room_rate,
               r.CommissionPercent AS comm_pct, r.Revenue AS revenue, r.USDRevenue AS usd_revenue,
               r.Country AS country, r.BookingAssociate AS associate,
               c.DateOnCheck AS date_on_check, c.Check_ AS check_number,
               c.FinalPayment AS final_payment, c.Advance AS is_advance,
               COALESCE(p.total_pickup,0) AS actual_pickup,
               COALESCE(all_pay.pay_total,0) AS total_pay_amount,
               COALESCE(all_pay.pay_count,0) AS total_pay_count
        FROM ChkRegNote c
        JOIN ReportPipeline r ON r.BookingId = c.BookingID
        LEFT JOIN (SELECT BookingID, SUM(ActualPickup) AS total_pickup FROM Pickup GROUP BY BookingID) p
            ON p.BookingID = r.BookingId
        LEFT JOIN (SELECT BookingID, COUNT(*) AS pay_count, SUM(COALESCE(FinalPayment,0)) AS pay_total
                   FROM ChkRegNote WHERE (Cancelled IS NULL OR Cancelled = 0) GROUP BY BookingID) all_pay
            ON all_pay.BookingID = c.BookingID
        WHERE DATE(c.DateOnCheck) BETWEEN ? AND ?
        AND (c.Cancelled IS NULL OR c.Cancelled = 0)
        AND (LOWER(COALESCE(r.BookingAssociate,'')) = 'kristin house'
             OR LOWER(COALESCE(r.BookingType,'')) NOT IN ('other services', 'other', 'conference management', 'cm'))
        ORDER BY c.DateOnCheck DESC
    '''
    raw = db.execute(query, (date_from, date_to)).fetchall()

    default_split  = get_commission_split()
    account_splits = get_account_splits()
    kristin_split  = get_kristin_split()
    kristin_cut    = get_kristin_cut()
    tolerance      = get_payment_tolerance()

    by_check_rows = []   # Tab 1 – all payments sorted by check #
    kristin_rows  = []   # Tab 2 – Kristin House only
    team_rows     = []   # Tab 3 – everyone else

    def _us(iso):
        try:
            return datetime.strptime(str(iso)[:10], '%Y-%m-%d').strftime('%m/%d/%Y')
        except Exception:
            return str(iso or '')[:10]

    for r in raw:
        r = dict(r)
        is_kristin = (r.get('associate') or '').strip().lower() == 'kristin house'
        final      = float(r['final_payment'] or 0)
        date_str   = _us(r.get('date_on_check')) if r.get('date_on_check') else ''

        by_check_rows.append({
            'Check #':        r['check_number'],
            'Date on Check':  date_str,
            'Associate':      r['associate'] or '',
            'Booking ID':     r['booking_id'],
            'Account':        r['account'] or '',
            'Event Name':     r['event_name'] or '',
            'Hotel':          r['hotel'] or '',
            'Incentive':      'Yes' if r['is_advance'] else '',
            'Final Payment':  final,
        })

        if is_kristin:
            split  = effective_split(r.get('associate'), r.get('account'), r.get('country'),
                                     account_splits, default_split, kristin_split, kristin_cut)
            pickup = float(r['actual_pickup'] or 0)
            try:
                actual = pickup * float(r['room_rate'] or 0) * float(r['comm_pct'] or 0) * split
                if actual == 0:
                    rev = float(r['usd_revenue'] or 0) or float(r['revenue'] or 0)
                    actual = rev * float(r['comm_pct'] or 0) * split
            except Exception:
                actual = 0
            total_paid  = float(r.get('total_pay_amount') or 0)
            prior_count = int(r.get('total_pay_count') or 0) - 1
            prior_total = total_paid - final if prior_count > 0 else 0
            kristin_rows.append({
                'Date on Check':     date_str,
                'Check #':           r['check_number'],
                'Booking ID':        r['booking_id'],
                'Account':           r['account'] or '',
                'Event Name':        r['event_name'] or '',
                'Hotel':             r['hotel'] or '',
                'Start Date':        _us(r['start_date']) if r['start_date'] else '',
                'Split %':           split_label(r.get('associate'), split, default_split, kristin_split, kristin_cut),
                'Final Payment':     final,
                'Actual Commission': round(actual, 2),
                'Difference':        round(final - actual, 2),
                'Prior Payments':    round(prior_total, 2) if prior_count > 0 else 0,
                'Grand Total Paid':  round(total_paid, 2),
                'Diff %':            round((total_paid - actual) / actual, 4) if actual else None,
                'Incentive':         'Yes' if r['is_advance'] else '',
            })
        else:
            team_rows.append({
                'Date on Check':  date_str,
                'Check #':        r['check_number'],
                'Associate':      r['associate'] or '',
                'Booking ID':     r['booking_id'],
                'Account':        r['account'] or '',
                'Event Name':     r['event_name'] or '',
                'Hotel':          r['hotel'] or '',
                'Start Date':     _us(r['start_date']) if r['start_date'] else '',
                'Incentive':      'Yes' if r['is_advance'] else '',
                'Final Payment':  final,
            })

    # Sort Tab 1 by check number then date; Tabs 2 & 3 by date paid
    by_check_rows.sort(key=lambda x: (str(x['Check #'] or ''), x['Date on Check']))
    kristin_rows.sort(key=lambda x: x['Date on Check'])
    team_rows.sort(key=lambda x: x['Date on Check'])

    from itertools import groupby as _groupby
    from openpyxl.styles import PatternFill, Font, Alignment

    HEADER_FILL = PatternFill('solid', start_color='1A3A5C')
    HEADER_FONT = Font(bold=True, color='FFFFFF')
    SUBTOT_FILL = PatternFill('solid', start_color='D9E1F2')
    GRAND_FILL  = PatternFill('solid', start_color='1A3A5C')
    GRAND_FONT  = Font(bold=True, color='FFFFFF')
    BOLD        = Font(bold=True)
    RIGHT       = Alignment(horizontal='right')
    CENTER      = Alignment(horizontal='center', wrap_text=True)

    def style_sheet(ws, currency_cols, col_widths):
        for cell in ws[1]:
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT
            cell.alignment = CENTER
        for i, w in enumerate(col_widths, 1):
            ws.column_dimensions[ws.cell(1, i).column_letter].width = w
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            for cell in row:
                if cell.column_letter in currency_cols:
                    cell.number_format = '$#,##0.00'
        # Grand totals row
        last = ws.max_row
        tr = last + 1
        first_curr_idx = min(ord(c) - ord('A') + 1 for c in currency_cols)
        lbl = ws.cell(tr, first_curr_idx - 1)
        lbl.value = 'Grand Total:'
        lbl.font = GRAND_FONT
        lbl.fill = GRAND_FILL
        lbl.alignment = RIGHT
        for cl in currency_cols:
            ci = ord(cl) - ord('A') + 1
            c = ws.cell(tr, ci)
            c.value = f'=SUM({cl}2:{cl}{last})'
            c.number_format = '$#,##0.00'
            c.font = GRAND_FONT
            c.fill = GRAND_FILL

    ORANGE_FILL = PatternFill('solid', start_color='FFD580')

    def write_grouped_sheet(ws, rows, group_key, label_prefix):
        """Write a sheet grouped by group_key with subtotals and 2 blank rows between groups."""
        headers    = ['Date on Check', 'Check #', 'Associate', 'Booking ID',
                      'Account', 'Event Name', 'Hotel', 'Start Date', 'Incentive', 'Final Payment']
        col_widths = [15, 14, 28, 14, 28, 35, 30, 13, 10, 15]
        pay_col    = 10

        for ci, h in enumerate(headers, 1):
            cell = ws.cell(1, ci, h)
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT
            cell.alignment = CENTER
        for i, w in enumerate(col_widths, 1):
            ws.column_dimensions[ws.cell(1, i).column_letter].width = w
        ws.freeze_panes = 'A2'

        cur_row     = 2
        grand_total = 0.0

        for grp_val, grp in _groupby(rows, key=lambda x: x[group_key]):
            grp_rows  = list(grp)
            grp_total = 0.0

            for r in grp_rows:
                is_incentive = r['Incentive'] == 'Yes'
                ws.cell(cur_row, 1, r['Date on Check'])
                ws.cell(cur_row, 2, r['Check #'])
                ws.cell(cur_row, 3, r['Associate'])
                ws.cell(cur_row, 4, r['Booking ID'])
                ws.cell(cur_row, 5, r['Account'])
                ws.cell(cur_row, 6, r['Event Name'])
                ws.cell(cur_row, 7, r['Hotel'])
                ws.cell(cur_row, 8, r['Start Date'])
                ws.cell(cur_row, 9, r['Incentive'])
                pay = ws.cell(cur_row, pay_col, r['Final Payment'])
                pay.number_format = '$#,##0.00'
                if is_incentive:
                    for ci in range(1, pay_col + 1):
                        ws.cell(cur_row, ci).fill = ORANGE_FILL
                grp_total += r['Final Payment']
                cur_row += 1

            lbl = ws.cell(cur_row, pay_col - 1, f'{label_prefix} {grp_val} Total:')
            lbl.font = BOLD
            lbl.fill = SUBTOT_FILL
            lbl.alignment = RIGHT
            tot = ws.cell(cur_row, pay_col, grp_total)
            tot.number_format = '$#,##0.00'
            tot.font = BOLD
            tot.fill = SUBTOT_FILL
            grand_total += grp_total
            cur_row += 3

        lbl = ws.cell(cur_row, pay_col - 1, 'Grand Total:')
        lbl.font = GRAND_FONT
        lbl.fill = GRAND_FILL
        lbl.alignment = RIGHT
        tot = ws.cell(cur_row, pay_col, grand_total)
        tot.number_format = '$#,##0.00'
        tot.font = GRAND_FONT
        tot.fill = GRAND_FILL

    def write_by_check_sheet(ws, rows):
        """Write Tab 1 with per-check subtotals, 2 blank rows between groups, frozen header."""
        headers    = ['Check #', 'Date on Check', 'Associate', 'Booking ID',
                      'Account', 'Event Name', 'Hotel', 'Incentive', 'Final Payment']
        col_widths = [14, 15, 28, 14, 28, 35, 30, 10, 15]

        for ci, h in enumerate(headers, 1):
            cell = ws.cell(1, ci, h)
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT
            cell.alignment = CENTER
        for i, w in enumerate(col_widths, 1):
            ws.column_dimensions[ws.cell(1, i).column_letter].width = w
        ws.freeze_panes = 'A2'

        cur_row     = 2
        grand_total = 0.0

        for check_num, grp in _groupby(rows, key=lambda x: x['Check #']):
            grp_rows    = list(grp)
            check_total = 0.0

            for r in grp_rows:
                is_incentive = r['Incentive'] == 'Yes'
                ws.cell(cur_row, 1, r['Check #'])
                ws.cell(cur_row, 2, r['Date on Check'])
                ws.cell(cur_row, 3, r['Associate'])
                ws.cell(cur_row, 4, r['Booking ID'])
                ws.cell(cur_row, 5, r['Account'])
                ws.cell(cur_row, 6, r['Event Name'])
                ws.cell(cur_row, 7, r['Hotel'])
                ws.cell(cur_row, 8, r['Incentive'])
                pay = ws.cell(cur_row, 9, r['Final Payment'])
                pay.number_format = '$#,##0.00'
                if is_incentive:
                    for ci in range(1, 10):
                        ws.cell(cur_row, ci).fill = ORANGE_FILL
                check_total += r['Final Payment']
                cur_row += 1

            lbl = ws.cell(cur_row, 8, f'Check {check_num} Total:')
            lbl.font = BOLD
            lbl.fill = SUBTOT_FILL
            lbl.alignment = RIGHT
            tot = ws.cell(cur_row, 9, check_total)
            tot.number_format = '$#,##0.00'
            tot.font = BOLD
            tot.fill = SUBTOT_FILL
            grand_total += check_total
            cur_row += 3

        lbl = ws.cell(cur_row, 8, 'Grand Total:')
        lbl.font = GRAND_FONT
        lbl.fill = GRAND_FILL
        lbl.alignment = RIGHT
        tot = ws.cell(cur_row, 9, grand_total)
        tot.number_format = '$#,##0.00'
        tot.font = GRAND_FONT
        tot.fill = GRAND_FILL

    # Tab 4 sort: by associate then date
    team_by_assoc = sorted(team_rows, key=lambda x: (x['Associate'], x['Date on Check']))

    # ── Summary data ──────────────────────────────────────────────────────────
    # Totals by check number (from all rows)
    check_totals = {}
    for r in by_check_rows:
        k = str(r['Check #'] or '')
        check_totals[k] = check_totals.get(k, 0.0) + r['Final Payment']
    check_totals = sorted(check_totals.items(), key=lambda x: x[0])

    # Totals by associate (Kristin + team)
    assoc_totals = {}
    for r in kristin_rows:
        assoc_totals['Kristin House'] = assoc_totals.get('Kristin House', 0.0) + r['Final Payment']
    for r in team_rows:
        a = r['Associate'] or 'Unknown'
        assoc_totals[a] = assoc_totals.get(a, 0.0) + r['Final Payment']
    assoc_totals = sorted(assoc_totals.items(), key=lambda x: x[0])

    def write_summary_sheet(ws):
        from openpyxl.styles import PatternFill, Font, Alignment
        from openpyxl.utils import get_column_letter

        TITLE_FONT   = Font(bold=True, size=14, color='1A3A5C')
        SECTION_FONT = Font(bold=True, size=11, color='FFFFFF')
        SECTION_FILL = PatternFill('solid', start_color='1A3A5C')
        SUBTOT_FONT  = Font(bold=True)
        CURRENCY_FMT = '$#,##0.00'
        CENTER       = Alignment(horizontal='center')
        RIGHT        = Alignment(horizontal='right')

        ws.column_dimensions['A'].width = 30
        ws.column_dimensions['B'].width = 18

        # Title
        ws.merge_cells('A1:B1')
        title = ws.cell(1, 1, 'Conference Planning Associates Inc.')
        title.font = TITLE_FONT
        title.alignment = CENTER

        ws.merge_cells('A2:B2')
        sub = ws.cell(2, 1, 'Check Register Report')
        sub.font = Font(bold=True, size=12, color='1A3A5C')
        sub.alignment = CENTER

        ws.merge_cells('A3:B3')
        dr = ws.cell(3, 1, f'{date_from}  —  {date_to}')
        dr.font = Font(size=10, color='555555')
        dr.alignment = CENTER

        cur = 5   # start data after blank row 4

        # ── By Check Number ────────────────────────────────────────────────
        ws.merge_cells(f'A{cur}:B{cur}')
        hdr = ws.cell(cur, 1, 'Total by Check Number')
        hdr.font = SECTION_FONT
        hdr.fill = SECTION_FILL
        hdr.alignment = CENTER
        cur += 1

        col_hdr_fill = PatternFill('solid', start_color='D9E1F2')
        for col, label in enumerate(['Check #', 'Total'], 1):
            c = ws.cell(cur, col, label)
            c.font = Font(bold=True)
            c.fill = col_hdr_fill
            c.alignment = CENTER
        cur += 1

        check_grand = 0.0
        for check_num, total in check_totals:
            ws.cell(cur, 1, check_num)
            amt = ws.cell(cur, 2, total)
            amt.number_format = CURRENCY_FMT
            check_grand += total
            cur += 1

        ws.cell(cur, 1, 'Grand Total:').font = SUBTOT_FONT
        ws.cell(cur, 1).alignment = RIGHT
        gtc = ws.cell(cur, 2, check_grand)
        gtc.number_format = CURRENCY_FMT
        gtc.font = SUBTOT_FONT
        gtc.fill = SECTION_FILL
        gtc.font = Font(bold=True, color='FFFFFF')
        ws.cell(cur, 1).fill = SECTION_FILL
        ws.cell(cur, 1).font = Font(bold=True, color='FFFFFF')
        cur += 2   # blank separator

        # ── By Associate ───────────────────────────────────────────────────
        ws.merge_cells(f'A{cur}:B{cur}')
        hdr = ws.cell(cur, 1, 'Total by Associate')
        hdr.font = SECTION_FONT
        hdr.fill = SECTION_FILL
        hdr.alignment = CENTER
        cur += 1

        for col, label in enumerate(['Associate', 'Total'], 1):
            c = ws.cell(cur, col, label)
            c.font = Font(bold=True)
            c.fill = col_hdr_fill
            c.alignment = CENTER
        cur += 1

        assoc_grand = 0.0
        for assoc, total in assoc_totals:
            ws.cell(cur, 1, assoc)
            amt = ws.cell(cur, 2, total)
            amt.number_format = CURRENCY_FMT
            assoc_grand += total
            cur += 1

        ws.cell(cur, 1, 'Grand Total:').fill = SECTION_FILL
        ws.cell(cur, 1).font = Font(bold=True, color='FFFFFF')
        ws.cell(cur, 1).alignment = RIGHT
        gta = ws.cell(cur, 2, assoc_grand)
        gta.number_format = CURRENCY_FMT
        gta.fill = SECTION_FILL
        gta.font = Font(bold=True, color='FFFFFF')

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Placeholders to control sheet order; Tabs 1 & 4 written manually
        pd.DataFrame(columns=['Summary']).to_excel(writer, index=False, sheet_name='Summary')
        pd.DataFrame(columns=['Check #','Date on Check','Associate','Booking ID',
                               'Account','Event Name','Hotel','Incentive','Final Payment']
                     ).to_excel(writer, index=False, sheet_name='By Check')
        pd.DataFrame(kristin_rows).to_excel(writer, index=False, sheet_name='Kristin House')
        pd.DataFrame(team_rows).to_excel(writer, index=False, sheet_name='Team')
        pd.DataFrame(columns=['Date on Check','Check #','Associate','Booking ID',
                               'Account','Event Name','Hotel','Start Date','Incentive','Final Payment']
                     ).to_excel(writer, index=False, sheet_name='Total by Associate')

        # ── Summary ───────────────────────────────────────────────────────────
        write_summary_sheet(writer.sheets['Summary'])

        # ── Tab 1: By Check ───────────────────────────────────────────────────
        write_by_check_sheet(writer.sheets['By Check'], by_check_rows)

        # ── Tab 2: Kristin House ──────────────────────────────────────────────
        # Cols A-N: Date Check# BookingID Account Event Hotel Start Split% Payment
        #           ActComm Diff Prior Grand Incentive
        # Cols: A=Date B=Check# C=BookingID D=Account E=Event F=Hotel G=Start H=Split%
        #       I=Payment J=ActComm K=Diff L=Prior M=Grand N=Diff% O=Incentive
        style_sheet(writer.sheets['Kristin House'],
                    ('I', 'J', 'K', 'L', 'M'),
                    [15, 14, 14, 28, 35, 30, 13, 10, 15, 17, 12, 14, 16, 10, 10])
        ws_k = writer.sheets['Kristin House']
        ws_k.freeze_panes = 'A2'
        # Format Diff % column (N) as percentage
        for row in ws_k.iter_rows(min_row=2, max_row=ws_k.max_row):
            if row[13].value is not None:
                row[13].number_format = '+0.0%;-0.0%;0.0%'
        red_fill    = PatternFill('solid', start_color='F8D7DA')
        yellow_fill = PatternFill('solid', start_color='FFF3CD')
        last_data   = ws_k.max_row - 1
        for row in ws_k.iter_rows(min_row=2, max_row=last_data):
            incentive   = row[14].value == 'Yes'   # col O
            fin         = row[8].value              # col I
            act         = row[9].value              # col J
            grand       = row[12].value or 0        # col M – Grand Total Paid
            start_date  = str(row[6].value or '')[:10]
            is_future   = start_date > today
            try:
                covered   = act and act > 0 and grand >= act * (1 - tolerance)
                underpaid = (act and act > 0 and fin is not None and fin < act
                             and (act - fin) / act > tolerance
                             and not covered)
            except Exception:
                underpaid = False
            if incentive:
                fill = ORANGE_FILL
            elif underpaid:
                fill = red_fill
            elif is_future:
                fill = yellow_fill
            else:
                fill = None
            if fill:
                for cell in row:
                    cell.fill = fill

        # ── Tab 3: Team (sorted by date) ──────────────────────────────────────
        # Cols A-J: Date Check# Associate BookingID Account Event Hotel Start Incentive Payment
        style_sheet(writer.sheets['Team'],
                    ('J',),
                    [15, 14, 28, 14, 28, 35, 30, 13, 10, 15])
        ws_t = writer.sheets['Team']
        ws_t.freeze_panes = 'A2'
        last_team = ws_t.max_row - 1
        for row in ws_t.iter_rows(min_row=2, max_row=last_team):
            if row[8].value == 'Yes':   # col I – Incentive
                for cell in row:
                    cell.fill = ORANGE_FILL

        # ── Tab 4: Total by Associate ──────────────────────────────────────────
        write_grouped_sheet(writer.sheets['Total by Associate'],
                            team_by_assoc, 'Associate', 'Associate')

    output.seek(0)
    filename = f'Payment_Report_{date_from}_to_{date_to}.xlsx'
    return send_file(output, download_name=filename, as_attachment=True,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

# ── Commission Report Export ──────────────────────────────────────────────────

@app.route('/reports/commission/export')
def report_commission_export():
    date_from = request.args.get('date_from', '').strip()
    date_to   = request.args.get('date_to', '').strip()
    who       = request.args.get('who', 'team')
    today     = datetime.now().strftime('%Y-%m-%d')
    if not date_from or not date_to:
        flash('Please run the report with a date range before exporting.', 'error')
        return redirect(url_for('report_commission'))

    db = get_db()
    query = '''
        SELECT r.BookingId AS booking_id, r.AccountName AS account, r.Customer AS hotel,
               r.BookedDate AS booked_date, r.StartDate AS start_date, r.EndDate AS end_date,
               r.RoomRate AS room_rate, r.TotalRoomNights AS total_room_nights,
               r.CommissionPercent AS comm_pct, r.Revenue AS revenue, r.USDRevenue AS usd_revenue,
               r.Country AS country, r.BookingAssociate AS associate,
               COALESCE(p.total_pickup,0) AS actual_pickup,
               CASE WHEN pay.pay_count > 0 THEN 'Paid / Future' ELSE 'Unpaid' END AS payment_status
        FROM ReportPipeline r
        LEFT JOIN (SELECT BookingID, SUM(ActualPickup) AS total_pickup FROM Pickup GROUP BY BookingID) p
            ON p.BookingID = r.BookingId
        LEFT JOIN (SELECT BookingID, COUNT(*) AS pay_count FROM ChkRegNote
                   WHERE (Cancelled IS NULL OR Cancelled=0) GROUP BY BookingID) pay
            ON pay.BookingID = r.BookingId
        WHERE DATE(r.StartDate) BETWEEN ? AND ?
        AND (pay.pay_count IS NULL OR pay.pay_count=0 OR DATE(r.StartDate) > ?)
        AND (r.BookingStatus IS NULL OR r.BookingStatus != 'Cancelled')
        AND NOT EXISTS (SELECT 1 FROM ChkRegNote c2 WHERE c2.BookingID = r.BookingId AND c2.Cancelled=1)
        AND (LOWER(COALESCE(r.BookingAssociate,'')) = 'kristin house'
             OR LOWER(COALESCE(r.BookingType,'')) NOT IN ('other services', 'other', 'conference management', 'cm'))
        ORDER BY r.StartDate
    '''
    all_rows = db.execute(query, (date_from, date_to, today)).fetchall()
    if who == 'kristin':
        rows = [r for r in all_rows if (r['associate'] or '').strip().lower() == 'kristin house']
    else:
        rows = [r for r in all_rows if (r['associate'] or '').strip().lower() != 'kristin house']

    default_split  = get_commission_split()
    account_splits = get_account_splits()
    kristin_split  = get_kristin_split()
    kristin_cut    = get_kristin_cut()
    export_rows     = []
    missing_pickups = set()   # 0-based data row indices with no pickup
    for r in rows:
        split = effective_split(r['associate'], r['account'], r['country'],
                                account_splits, default_split, kristin_split, kristin_cut)
        try:
            rev = float(r['usd_revenue'] or 0) or float(r['revenue'] or 0)
            est = rev * float(r['comm_pct'] or 0) * split
        except Exception:
            est = 0
        pickup = float(r['actual_pickup'] or 0)
        try:
            actual = pickup * float(r['room_rate'] or 0) * float(r['comm_pct'] or 0) * split
        except Exception:
            actual = 0
        if pickup == 0:
            missing_pickups.add(len(export_rows))
        def _us2(iso):
            try:
                return datetime.strptime(str(iso)[:10], '%Y-%m-%d').strftime('%m/%d/%Y')
            except Exception:
                return str(iso or '')
        export_rows.append({
            'Booking ID':        r['booking_id'],
            'Account':           r['account'],
            'Hotel':             r['hotel'],
            'Booked Date':       _us2(r['booked_date']) if r['booked_date'] else '',
            'Start Date':        _us2(r['start_date']) if r['start_date'] else '',
            'End Date':          _us2(r['end_date']) if r['end_date'] else '',
            'Room Rate':         r['room_rate'],
            'Contracted Nights': int(r['total_room_nights']) if r['total_room_nights'] else 0,
            'Pickup':            int(r['actual_pickup']) if r['actual_pickup'] else 0,
            'Est. Commission':   round(est, 2),
            'Actual Commission': round(actual, 2),
            'Status':            r['payment_status'],
        })

    df = pd.DataFrame(export_rows)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Missing Commission')
        ws = writer.sheets['Missing Commission']
        ws.freeze_panes = 'A2'
        col_widths = [14,30,35,14,14,14,12,14,10,16,18,14]
        for i, width in enumerate(col_widths, 1):
            ws.column_dimensions[ws.cell(1,i).column_letter].width = width
        from openpyxl.styles import PatternFill, Font, Alignment
        header_fill = PatternFill('solid', start_color='1A3A5C')
        header_font = Font(bold=True, color='FFFFFF')
        red_fill    = PatternFill('solid', start_color='F8D7DA')
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', wrap_text=True)
        for idx, row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row)):
            for cell in row:
                if cell.column_letter in ('G','J','K'):
                    cell.number_format = '$#,##0.00'
                if cell.column_letter in ('H','I'):
                    cell.number_format = '#,##0'
            if idx in missing_pickups:
                for cell in row:
                    cell.fill = red_fill
        total_row = ws.max_row + 1
        last_data  = total_row - 1
        ws.cell(total_row, 9).value = 'Totals:'
        ws.cell(total_row, 9).font = Font(bold=True)
        ws.cell(total_row, 9).alignment = Alignment(horizontal='right')
        ws.cell(total_row,10).value = f'=SUM(J2:J{last_data})'
        ws.cell(total_row,10).number_format = '$#,##0.00'
        ws.cell(total_row,10).font = Font(bold=True)
        ws.cell(total_row,11).value = f'=SUM(K2:K{last_data})'
        ws.cell(total_row,11).number_format = '$#,##0.00'
        ws.cell(total_row,11).font = Font(bold=True)

    output.seek(0)
    who_label = 'Kristin_House' if who == 'kristin' else 'Team'
    filename = f'Missing_Commission_{who_label}_{date_from}_to_{date_to}.xlsx'
    return send_file(output, download_name=filename, as_attachment=True,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

# ── Customer Summary Report ────────────────────────────────────────────────────

from cs_report_utils import (
    _chain_color, _query_customer_summary, _query_meeting_summary,
    _build_word_doc, _build_pptx, build_city_map_data
)

@app.route('/reports/customer-summary', methods=['GET', 'POST'])
def report_customer_summary():
    user = get_current_user()
    if not has_permission(user, 'reports_customer_summary'):
        flash('You do not have access to the Customer Summary report.', 'error')
        return redirect(url_for('index'))
    export = request.args.get('export', '').lower()
    conn   = get_db()

    customers = [r[0] for r in conn.execute(
        '''SELECT DISTINCT AccountName FROM ReportPipeline
           WHERE AccountName IS NOT NULL
             AND substr(StartDate,1,10) >= '2025-01-01'
           ORDER BY AccountName'''
    ).fetchall()]

    # Team member dropdown — filtered to 2025+, Kristin House pinned first
    raw_team_members = [r[0] for r in conn.execute(
        '''SELECT DISTINCT BookingAssociate FROM ReportPipeline
           WHERE BookingAssociate IS NOT NULL AND BookingAssociate != ''
             AND BookingAssociate NOT IN ('Unknown','Unkown','Housing Registration')
             AND substr(StartDate,1,10) >= '2025-01-01'
           ORDER BY BookingAssociate'''
    ).fetchall()]
    team_members = []
    if 'Kristin House' in raw_team_members:
        team_members.append('Kristin House')
    team_members += [tm for tm in raw_team_members if tm != 'Kristin House']

    # Build team-member → customers map for dynamic JS filtering
    tm_cust_rows = conn.execute(
        '''SELECT DISTINCT BookingAssociate, AccountName FROM ReportPipeline
           WHERE AccountName IS NOT NULL AND BookingAssociate IS NOT NULL
             AND BookingAssociate != ''
             AND BookingAssociate NOT IN ('Unknown','Unkown','Housing Registration')
             AND substr(StartDate,1,10) >= '2025-01-01'
           ORDER BY BookingAssociate, AccountName'''
    ).fetchall()
    tm_customers = {}
    for tm, acct in tm_cust_rows:
        tm_customers.setdefault(tm, []).append(acct)

    today = datetime.today()
    default_start = f'{today.year}-01-01'
    default_end   = f'{today.year}-12-31'

    if request.method == 'POST':
        selected_customers = request.form.getlist('customer')
    else:
        selected_customers = request.args.getlist('customer')
    selected_customers = [c.strip() for c in selected_customers if c.strip()]

    date_from   = (request.form.get('date_from')   or request.args.get('date_from',   default_start)).strip()
    date_to     = (request.form.get('date_to')     or request.args.get('date_to',     default_end)).strip()
    team_member = (request.form.get('team_member') or request.args.get('team_member', '')).strip()
    submitted   = request.method == 'POST' or bool(export)

    brand_rows, chain_groups, grand_total, grand_bookings = [], {}, 0.0, 0
    city_meetings = {}

    if request.method == 'POST' or export:
        brand_rows, chain_groups, grand_total, grand_bookings = _query_customer_summary(
            conn, selected_customers, date_from, date_to,
            customer_col='AccountName', status_col='BookingStatus',
            date_col='StartDate', contract_col='ContractedAmount',
            revenue_col='Revenue', chain_col='Chain', brand_col='Brand',
            city_col='City', state_col='State', cast_contract=False,
            team_member=team_member, team_member_col='BookingAssociate')
        city_meetings = _query_meeting_summary(
            conn, selected_customers, date_from, date_to,
            customer_col='AccountName', status_col='BookingStatus',
            date_col='StartDate', contract_col='ContractedAmount',
            revenue_col='Revenue', event_col='EventName',
            booking_col='BookingId', hotel_col='Customer',
            city_col='City', state_col='State', cast_contract=False,
            team_member=team_member, team_member_col='BookingAssociate')
        submitted = True

    # Build export filename/subtitle including team member when filtered
    tm_slug = f'_{team_member.replace(" ","_")}' if team_member else ''
    if len(selected_customers) == 1:
        cust_slug = selected_customers[0].replace(' ', '_')
    elif selected_customers:
        cust_slug = f'{len(selected_customers)}_Customers'
    else:
        cust_slug = 'All'

    city_map_data = build_city_map_data(city_meetings) if city_meetings else []

    if export == 'word' and brand_rows:
        buf = _build_word_doc(selected_customers, date_from, date_to, brand_rows, chain_groups,
                              grand_total, grand_bookings, city_meetings,
                              team_member=team_member)
        fname = f'Customer_Summary_{cust_slug}{tm_slug}_{date_from}_{date_to}.docx'
        return send_file(buf, download_name=fname, as_attachment=True,
                         mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')

    if export == 'pptx' and brand_rows:
        buf = _build_pptx(selected_customers, date_from, date_to, brand_rows, chain_groups,
                          grand_total, grand_bookings, city_meetings,
                          team_member=team_member, city_map_data=city_map_data)
        fname = f'Customer_Summary_{cust_slug}{tm_slug}_{date_from}_{date_to}.pptx'
        return send_file(buf, download_name=fname, as_attachment=True,
                         mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation')

    brand_labels = [r['chain'] for r in brand_rows]
    brand_values = [r['revenue'] or 0 for r in brand_rows]
    chain_colors = [_chain_color(l, i) for i, l in enumerate(brand_labels)]

    city_rev = {}
    for data in chain_groups.values():
        for c in data['cities']:
            key = f"{c['city']}, {c['state']}" if c['state'] else c['city']
            city_rev[key] = city_rev.get(key, 0) + c['revenue']
    top_cities = sorted(city_rev.items(), key=lambda x: x[1], reverse=True)[:14]
    city_labels = [x[0] for x in top_cities]
    city_values = [x[1] for x in top_cities]
    city_count  = len(city_rev)

    return render_template('report_customer_summary.html',
        customers=customers, selected_customers=selected_customers,
        date_from=date_from, date_to=date_to, submitted=submitted,
        brand_rows=brand_rows, chain_groups=chain_groups,
        grand_total=grand_total, grand_bookings=grand_bookings,
        brand_labels=brand_labels, brand_values=brand_values,
        city_labels=city_labels, city_values=city_values,
        chain_colors=chain_colors, city_count=city_count,
        city_meetings=city_meetings, city_map_data=city_map_data,
        team_members=team_members, team_member=team_member,
        tm_customers=tm_customers,
    )

# ── Settings ──────────────────────────────────────────────────────────────────

@app.route('/settings', methods=['GET', 'POST'])
def settings():
    db = get_db()
    if request.method == 'POST':
        action = request.form.get('action')
        if action == 'save_default':
            split     = request.form.get('commission_split', '').strip()
            tolerance = request.form.get('payment_tolerance', '').strip()
            try:
                val = float(split.replace('%',''))
                if val > 1: val = val / 100
                tval = float(tolerance.replace('%',''))
                if tval > 1: tval = tval / 100
                db.execute('UPDATE Settings SET value=? WHERE key="commission_split"', (str(val),))
                db.execute('UPDATE Settings SET value=? WHERE key="payment_tolerance"', (str(tval),))
                db.commit()
                flash(f'Settings saved — split: {val*100:.1f}%, tolerance: {tval*100:.1f}%.', 'success')
            except Exception:
                flash('Invalid value — please enter numbers like 60 or 1.', 'error')
        elif action == 'save_kristin':
            ksplit = request.form.get('kristin_split', '').strip()
            kcut   = request.form.get('kristin_cut', '').strip()
            try:
                kval = float(ksplit.replace('%',''))
                if kval > 1: kval = kval / 100
                cval = float(kcut.replace('%',''))
                if cval > 1: cval = cval / 100
                db.execute('INSERT OR REPLACE INTO Settings (key, value) VALUES ("kristin_split", ?)', (str(kval),))
                db.execute('INSERT OR REPLACE INTO Settings (key, value) VALUES ("kristin_cut", ?)', (str(cval),))
                db.commit()
                flash(f'Kristin House settings saved — Kristin split: {kval*100:.0f}%, team cut: {cval*100:.0f}%.', 'success')
            except Exception:
                flash('Invalid value — please enter numbers like 70 or 10.', 'error')
        elif action == 'add_account':
            account   = request.form.get('account_name', '').strip()
            split     = request.form.get('account_split', '').strip()
            countries = request.form.get('account_countries', '').strip()
            try:
                val = float(split.replace('%',''))
                if val > 1: val = val / 100
                db.execute('INSERT INTO AccountSplits (account_name, split_rate, countries) VALUES (?,?,?)',
                           (account, val, countries.upper()))
                db.commit()
                note = f' ({countries.upper()})' if countries else ' (all countries)'
                flash(f'Account override added: {account}{note} → {val*100:.1f}%.', 'success')
            except Exception:
                flash('Invalid entry — check account name and split value.', 'error')
        elif action == 'delete_account':
            acct_id = request.form.get('acct_id')
            db.execute('DELETE FROM AccountSplits WHERE id = ?', (acct_id,))
            db.commit()
            flash('Account override removed.', 'success')
        elif action == 'save_date_format':
            fmt = request.form.get('date_format', '%m/%d/%Y')
            known = {o[0] for o in DATE_FORMAT_OPTIONS}
            if fmt not in known:
                flash('Unknown date format selected.', 'error')
            else:
                db.execute('INSERT OR REPLACE INTO Settings (key, value) VALUES ("date_format", ?)', (fmt,))
                db.commit()
                label = next((o[1] for o in DATE_FORMAT_OPTIONS if o[0] == fmt), fmt)
                flash(f'Date format updated to {label}.', 'success')
        return redirect(url_for('settings'))

    current       = get_commission_split()
    tolerance     = get_payment_tolerance()
    kristin_split = get_kristin_split()
    kristin_cut   = get_kristin_cut()
    current_date_fmt = get_date_format()
    today_preview    = datetime.today().strftime(current_date_fmt)
    overrides = db.execute('SELECT id, account_name, split_rate, countries FROM AccountSplits ORDER BY account_name').fetchall()
    accounts  = [r[0] for r in db.execute('SELECT DISTINCT AccountName FROM ReportPipeline WHERE AccountName IS NOT NULL ORDER BY AccountName').fetchall()]
    return render_template('settings.html', commission_split=current, tolerance=tolerance,
                           kristin_split=kristin_split, kristin_cut=kristin_cut,
                           overrides=overrides, accounts=accounts,
                           date_format_options=DATE_FORMAT_OPTIONS,
                           current_date_fmt=current_date_fmt,
                           today_preview=today_preview)

# ── Summary ───────────────────────────────────────────────────────────────────

@app.route('/summary')
def summary():
    db = get_db()
    by_associate = db.execute('''
        SELECT BookingAssociate,
               COUNT(*) as Bookings,
               SUM(CASE WHEN BookingStatus='Definite' THEN 1 ELSE 0 END) as Definite,
               SUM(USDCommissionableAmount) as TotalCommission,
               SUM(USDRevenue) as TotalRevenue
        FROM ReportPipeline WHERE BookingAssociate IS NOT NULL
        GROUP BY BookingAssociate ORDER BY TotalCommission DESC
    ''').fetchall()
    by_status = db.execute('''
        SELECT BookingStatus, COUNT(*) as Count,
               SUM(USDCommissionableAmount) as TotalCommission,
               SUM(USDRevenue) as TotalRevenue
        FROM ReportPipeline GROUP BY BookingStatus ORDER BY Count DESC
    ''').fetchall()
    by_brand = db.execute('''
        SELECT Brand, COUNT(*) as Count, SUM(USDCommissionableAmount) as TotalCommission
        FROM ReportPipeline WHERE Brand IS NOT NULL
        GROUP BY Brand ORDER BY Count DESC LIMIT 20
    ''').fetchall()
    totals = db.execute('''
        SELECT COUNT(*) as Total,
               SUM(CASE WHEN BookingStatus='Definite' THEN 1 ELSE 0 END) as Definite,
               SUM(USDRevenue) as TotalRevenue,
               SUM(USDCommissionableAmount) as TotalCommission
        FROM ReportPipeline
    ''').fetchone()
    return render_template('summary.html', by_associate=by_associate, by_status=by_status,
                           by_brand=by_brand, totals=totals)

# ── Import Pickup Reports (HHR) ───────────────────────────────────────────────

def get_usd_exchange_rate(currency_code):
    """Fetch live USD exchange rate for a given 3-letter currency code."""
    import urllib.request, json as _json2
    currency_code = currency_code.upper()
    if currency_code == 'USD':
        return 1.0, 'USD'
    try:
        url = f'https://api.exchangerate-api.com/v4/latest/{currency_code}'
        with urllib.request.urlopen(url, timeout=5) as resp:
            data = _json2.loads(resp.read())
        rate = data['rates'].get('USD')
        if rate:
            return float(rate), 'live'
    except Exception:
        pass
    return None, None


CURRENCY_LABEL_MAP = {
    'usd': 'USD', 'us dollar': 'USD', 'dollar': 'USD',
    'euro': 'EUR', 'eur': 'EUR',
    'pound': 'GBP', 'gbp': 'GBP', 'british pound': 'GBP', 'sterling': 'GBP',
    'canadian': 'CAD', 'cad': 'CAD', 'canadian dollar': 'CAD',
    'australian': 'AUD', 'aud': 'AUD', 'australian dollar': 'AUD',
    'mexican': 'MXN', 'mxn': 'MXN', 'peso': 'MXN',
    'swiss': 'CHF', 'chf': 'CHF', 'franc': 'CHF',
    'japanese': 'JPY', 'jpy': 'JPY', 'yen': 'JPY',
}

REVENUE_ROW_KEYWORDS = [
    'TOTAL ROOMS COMMISSION DUE',
    'TOTAL ACCOMMODATION',
]


def parse_pickup_report(filepath):
    """Parse a TeamHouse Housing History Report xlsx and return extracted fields."""
    df = pd.read_excel(filepath, header=None)

    result = {
        'booking_id':       None,
        'hotel':            None,
        'organization':     None,
        'event_name':       None,
        'comm_pct':         None,
        'actual_pickup':    None,
        'pickup_by_night':  None,   # dict {"YYYY-MM-DD": rooms}
        'total_revenue':    None,
        'avg_rate':         None,
        'currency_code':    'USD',
        'currency_label':   'USD',
        'original_revenue': None,
        'exchange_rate':    None,
        'filename':         os.path.basename(filepath),
        'error':            None,
    }

    try:
        import re as _re2
        # Scan first 10 rows for Booking ID, Comm %, and Currency
        for row_idx in range(min(10, len(df))):
            row_raw = list(df.iloc[row_idx].values)
            row_str = [str(v) for v in row_raw]
            for i, cell in enumerate(row_str):
                if 'Booking' in cell and not result['booking_id']:
                    m = _re2.search(r'\d{5,7}', cell)
                    if m:
                        result['booking_id'] = m.group(0)
                    else:
                        for offset in (1, 2, 3):
                            if i + offset < len(row_raw):
                                val = row_raw[i + offset]
                                if val is None or (hasattr(val, '__float__') and str(val) == 'nan'):
                                    continue
                                try:
                                    bid = int(float(str(val)))
                                    if 10000 <= bid <= 9999999:
                                        result['booking_id'] = str(bid)
                                        break
                                except Exception:
                                    break
                if 'Comm %' in cell and not result['comm_pct']:
                    for offset in (1, 2, 3):
                        if i + offset < len(row_raw):
                            try:
                                v = float(row_raw[i + offset])
                                if not pd.isna(v) and v > 0:
                                    result['comm_pct'] = v
                                    break
                            except Exception:
                                pass
                if 'Currency:' in cell:
                    m = _re2.search(r'Currency:\s*(.+)', cell, _re2.IGNORECASE)
                    if m:
                        label = m.group(1).strip()
                        result['currency_label'] = label
                        iso = CURRENCY_LABEL_MAP.get(label.lower())
                        if iso:
                            result['currency_code'] = iso
            if result['booking_id'] and result['comm_pct']:
                break

        # Fallback: scan first 10 rows for a bare 6-7 digit integer
        if not result['booking_id']:
            for row_idx in range(min(10, len(df))):
                for val in df.iloc[row_idx].values:
                    if val is None or (hasattr(val, '__float__') and pd.isna(float(val)) if hasattr(val, '__float__') else False):
                        continue
                    try:
                        candidate = int(float(str(val)))
                        if 100000 <= candidate <= 9999999:
                            result['booking_id'] = str(candidate)
                            break
                    except Exception:
                        pass
                if result['booking_id']:
                    break

        # Row 1: Organization, Row 2: Hotel, Row 3: Event name
        for row_idx, key in [(1, 'organization'), (2, 'hotel'), (3, 'event_name')]:
            row_vals = df.iloc[row_idx].dropna().tolist()
            if len(row_vals) >= 2:
                result[key] = str(row_vals[1]).strip()

        # Find DATE header row → map column index to ISO date
        date_col_map = {}
        for idx, row in df.iterrows():
            row_vals = list(row.values)
            if str(row_vals[0]).strip().upper() == 'DATE':
                for ci, val in enumerate(row_vals):
                    if hasattr(val, 'strftime'):
                        date_col_map[ci] = val.strftime('%Y-%m-%d')
                break

        # Find FINAL TOTAL PICKUP row → total rooms + night-by-night breakdown
        for idx, row in df.iterrows():
            row_str_joined = ' '.join(str(v) for v in row.values if pd.notna(v))
            if 'FINAL TOTAL PICKUP' in row_str_joined.upper():
                nums = []
                pickup_by_night = {}
                for ci, v in enumerate(row.values):
                    try:
                        f = float(v)
                        if not pd.isna(f):
                            nums.append(f)
                            if ci in date_col_map and f > 0:
                                pickup_by_night[date_col_map[ci]] = int(f)
                    except Exception:
                        pass
                if nums:
                    result['actual_pickup'] = float(nums[-1])
                if pickup_by_night:
                    result['pickup_by_night'] = pickup_by_night
                break

        # Find TOTAL ROOMS / TOTAL ACCOMMODATION COMMISSION DUE row → revenue and avg rate
        for idx, row in df.iterrows():
            row_str_joined = ' '.join(str(v) for v in row.values if pd.notna(v)).upper()
            if any(kw in row_str_joined for kw in REVENUE_ROW_KEYWORDS) and 'COMMISSION DUE' in row_str_joined:
                nums = []
                found_label = False
                for v in row.values:
                    v_up = str(v).upper() if pd.notna(v) else ''
                    if pd.notna(v) and any(kw in v_up for kw in REVENUE_ROW_KEYWORDS):
                        found_label = True
                        continue
                    if found_label and pd.notna(v):
                        try:
                            nums.append(float(v))
                        except Exception:
                            pass
                if len(nums) >= 2:
                    result['total_revenue'] = nums[1]
                elif len(nums) == 1 and nums[0] > 10000:
                    # Some formats put only the commission total here; skip — use fallback below
                    pass
                if len(nums) >= 3:
                    result['avg_rate'] = nums[2]
                break

        # Fallback: scan for "Total Actualized Pickup Revenue" (Renaissance-style HHR format)
        # This row contains the actual room revenue when the commission row doesn't have it.
        if not result['total_revenue']:
            for idx, row in df.iterrows():
                row_str_joined = ' '.join(str(v) for v in row.values if pd.notna(v)).upper()
                if 'TOTAL ACTUALIZED PICKUP REVENUE' in row_str_joined or \
                   ('ACTUALIZED' in row_str_joined and 'REVENUE' in row_str_joined and 'PICKUP' in row_str_joined):
                    rev_nums = []
                    for v in row.values:
                        if pd.notna(v):
                            try:
                                f = float(v)
                                if not pd.isna(f) and f > 0:
                                    rev_nums.append(f)
                            except Exception:
                                pass
                    if rev_nums:
                        # Take the largest number as the revenue figure
                        result['total_revenue'] = max(rev_nums)
                    break

        # Fallback: derive avg_rate from revenue ÷ actual_pickup when not directly found
        if result['total_revenue'] and result['actual_pickup'] and not result['avg_rate']:
            try:
                result['avg_rate'] = round(float(result['total_revenue']) / float(result['actual_pickup']), 2)
            except Exception:
                pass

        # Currency conversion: if not USD, fetch live rate and convert revenue to USD
        if result['currency_code'] != 'USD' and result['total_revenue']:
            rate, source = get_usd_exchange_rate(result['currency_code'])
            if rate:
                result['original_revenue'] = result['total_revenue']
                result['exchange_rate']    = rate
                result['total_revenue']    = round(result['total_revenue'] * rate, 2)
                if result['avg_rate']:
                    result['avg_rate'] = round(result['avg_rate'] * rate, 2)
            else:
                result['error'] = (
                    f"Revenue is in {result['currency_label']} but live exchange rate "
                    f"could not be fetched. Please convert manually."
                )

        if not result['booking_id']:
            result['error'] = 'Could not find Booking ID in file'

    except Exception as e:
        result['error'] = str(e)

    return result


@app.route('/import/hhr', methods=['GET', 'POST'])
def import_hhr():
    user = get_current_user()
    if not has_permission(user, 'import_hhr'):
        flash('You do not have access to this import tool.', 'error')
        return redirect(url_for('index'))
    if request.method == 'POST':
        files = request.files.getlist('files')
        if not files or not any(f.filename for f in files):
            flash('No files selected.', 'error')
            return redirect(url_for('import_hhr'))

        db = get_db()
        results = []

        for file in files:
            if not file.filename:
                continue
            ext = file.filename.rsplit('.', 1)[-1].lower()
            if ext not in ('xls', 'xlsx'):
                results.append({'filename': file.filename, 'status': 'skipped',
                                'reason': 'Not an Excel file'})
                continue

            tmp_path = f'/tmp/cpainc_pickup_{file.filename}'
            file.save(tmp_path)
            file_bytes = open(tmp_path, 'rb').read()

            parsed = parse_pickup_report(tmp_path)

            if parsed['error']:
                results.append({'filename': file.filename, 'status': 'error',
                                'reason': parsed['error']})
                continue

            # Reject files with no pickup or revenue data
            has_pickup  = parsed['actual_pickup'] and float(parsed['actual_pickup']) > 0
            has_revenue = parsed['total_revenue'] and float(parsed['total_revenue']) > 0
            if not has_pickup and not has_revenue:
                results.append({'filename': file.filename, 'status': 'error',
                                'booking_id': str(parsed['booking_id']).strip().split('.')[0] if parsed['booking_id'] else None,
                                'reason': 'File contains no pickup rooms or revenue — nothing to import'})
                continue
            if not has_pickup:
                results.append({'filename': file.filename, 'status': 'error',
                                'booking_id': str(parsed['booking_id']).strip().split('.')[0] if parsed['booking_id'] else None,
                                'reason': 'File contains no actual pickup room nights — nothing to import'})
                continue
            if not has_revenue:
                results.append({'filename': file.filename, 'status': 'error',
                                'booking_id': str(parsed['booking_id']).strip().split('.')[0] if parsed['booking_id'] else None,
                                'reason': 'File contains no revenue total — nothing to import'})
                continue

            bid = str(parsed['booking_id']).strip().split('.')[0]

            # Check if booking exists
            booking = db.execute(
                'SELECT BookingId, EventName, Customer FROM ReportPipeline WHERE BookingId = ?', (bid,)
            ).fetchone()

            if not booking:
                from urllib.parse import urlencode
                qs = urlencode({
                    'BookingId':         bid,
                    'EventName':         parsed.get('event_name') or '',
                    'Customer':          parsed.get('hotel') or '',
                    'AccountName':       parsed.get('organization') or '',
                    'CommissionPercent': parsed.get('comm_pct') or '',
                })
                results.append({'filename': file.filename, 'status': 'no_booking',
                                'reason': f'Booking ID {bid} not found in database — click Create Booking to add it, then re-import.',
                                'booking_id': bid, 'create_url': f'/booking/new?{qs}'})
                continue

            # Create/update Pickup record
            existing = db.execute('SELECT ID FROM Pickup WHERE BookingID = ?', (bid,)).fetchone()
            if existing:
                db.execute('''UPDATE Pickup SET ActualPickup=?, TotalRevenue=?, Brand=?
                    WHERE ID=?''',
                    (parsed['actual_pickup'], parsed['total_revenue'],
                     parsed['hotel'] or '', existing['ID']))
                action = 'updated'
            else:
                db.execute('''INSERT INTO Pickup (BookingID, ActualPickup, TotalRevenue, Brand)
                    VALUES (?,?,?,?)''',
                    (bid, parsed['actual_pickup'], parsed['total_revenue'], parsed['hotel'] or ''))
                action = 'imported'

            # Update RoomRate on the booking if avg_rate available
            if parsed['avg_rate']:
                db.execute('UPDATE ReportPipeline SET RoomRate=? WHERE BookingId=?',
                           (parsed['avg_rate'], bid))

            # Archive the raw file
            db.execute(
                'INSERT INTO housing_history_files (booking_id, filename, file_data) VALUES (?,?,?)',
                (bid, file.filename, file_bytes)
            )

            # Create a "Final History" pickup_weekly entry if pickup_config exists
            final_history_result = None
            if parsed.get('pickup_by_night'):
                pc = db.execute(
                    "SELECT * FROM pickup_config WHERE booking_id=? AND status='active'", (bid,)
                ).fetchone()
                if pc:
                    existing_fh = db.execute(
                        "SELECT id FROM pickup_weekly WHERE config_id=? AND label='Final History'",
                        (pc['id'],)
                    ).fetchone()
                    if not existing_fh:
                        pbn       = parsed['pickup_by_night']
                        total_rms = sum(pbn.values())
                        block     = json.loads(pc['contracted_block'] or '{}')
                        total_blk = sum(block.values())
                        pct_blk   = round(total_rms / total_blk * 100, 1) if total_blk else None
                        atr_pct   = pc['attrition_pct']
                        atr_floor = round(total_blk * float(atr_pct), 1) if atr_pct and total_blk else None
                        pct_atr   = round(total_rms / atr_floor * 100, 1) if atr_floor else None
                        prev = db.execute(
                            "SELECT total_rooms FROM pickup_weekly "
                            "WHERE config_id=? AND label IS NULL "
                            "ORDER BY report_date DESC, id DESC LIMIT 1",
                            (pc['id'],)
                        ).fetchone()
                        change = (total_rms - prev['total_rooms']) if prev and prev['total_rooms'] else None
                        db.execute(
                            '''INSERT INTO pickup_weekly
                               (config_id, report_date, pickup_by_night, total_rooms,
                                change_from_last, pct_of_block, pct_of_attrition, label)
                               VALUES (?,?,?,?,?,?,?,?)''',
                            (pc['id'], datetime.now().strftime('%Y-%m-%d'),
                             json.dumps(pbn), total_rms, change, pct_blk, pct_atr, 'Final History')
                        )
                        final_history_result = {
                            'config_id': pc['id'],
                            'event':     pc['event_name'] or pc['organization'],
                            'total':     total_rms,
                            'pct_blk':   pct_blk,
                        }
                    else:
                        final_history_result = {'skipped': True}

            results.append({
                'filename':         file.filename,
                'status':           action,
                'booking_id':       bid,
                'event':            booking['EventName'],
                'hotel':            booking['Customer'],
                'pickup':           parsed['actual_pickup'],
                'revenue':          parsed['total_revenue'],
                'avg_rate':         parsed['avg_rate'],
                'currency_code':    parsed['currency_code'],
                'currency_label':   parsed['currency_label'],
                'original_revenue': parsed['original_revenue'],
                'exchange_rate':    parsed['exchange_rate'],
                'final_history':    final_history_result,
            })

        db.commit()
        return render_template('import_hhr_result.html', results=results)

    return render_template('import_hhr.html')


# ── Pickup Tracking Module ─────────────────────────────────────────────────────

def ensure_pickup_tables():
    db = get_db()
    db.executescript('''
        CREATE TABLE IF NOT EXISTS pickup_config (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            booking_id TEXT,
            tab_name TEXT,
            organization TEXT NOT NULL,
            event_name TEXT,
            hotel TEXT,
            hotel_contact TEXT,
            hotel_contact_email TEXT,
            group_contact TEXT,
            group_contact_email TEXT,
            cutoff_date TEXT,
            attrition_pct REAL,
            contracted_block TEXT NOT NULL DEFAULT '{}',
            contracted_rate REAL,
            shoulder_pre INTEGER DEFAULT 3,
            shoulder_post INTEGER DEFAULT 3,
            hotel_booking_link TEXT,
            status TEXT DEFAULT 'active',
            notes TEXT,
            created_at TEXT DEFAULT (datetime('now'))
        );
        CREATE TABLE IF NOT EXISTS pickup_weekly (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            config_id INTEGER NOT NULL REFERENCES pickup_config(id) ON DELETE CASCADE,
            report_date TEXT NOT NULL,
            pickup_by_night TEXT NOT NULL DEFAULT '{}',
            total_rooms INTEGER,
            change_from_last INTEGER,
            pct_of_block REAL,
            pct_of_attrition REAL,
            ota_rate REAL,
            label TEXT,
            notes TEXT,
            created_at TEXT DEFAULT (datetime('now'))
        );
        CREATE TABLE IF NOT EXISTS pickup_rooming_list (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            config_id INTEGER NOT NULL REFERENCES pickup_config(id) ON DELETE CASCADE,
            weekly_id INTEGER REFERENCES pickup_weekly(id),
            upload_date TEXT DEFAULT (datetime('now')),
            filename TEXT,
            file_data BLOB,
            total_guests INTEGER,
            nights_by_date TEXT,
            reconciliation_status TEXT DEFAULT 'pending',
            discrepancy_notes TEXT,
            guests_json TEXT
        );
        CREATE TABLE IF NOT EXISTS pickup_contact_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            config_id INTEGER NOT NULL REFERENCES pickup_config(id) ON DELETE CASCADE,
            contact_date TEXT NOT NULL,
            contact_type TEXT NOT NULL DEFAULT 'email_sent',
            notes TEXT,
            created_at TEXT DEFAULT (datetime('now'))
        );
        CREATE TABLE IF NOT EXISTS booking_contract (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            booking_id TEXT NOT NULL,
            filename TEXT NOT NULL,
            file_data BLOB NOT NULL,
            upload_date TEXT DEFAULT (date('now'))
        );
        CREATE TABLE IF NOT EXISTS housing_history_files (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            booking_id TEXT NOT NULL,
            filename TEXT NOT NULL,
            file_data BLOB NOT NULL,
            upload_date TEXT DEFAULT (datetime('now'))
        );
        CREATE TABLE IF NOT EXISTS client_contacts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            email TEXT NOT NULL,
            organization TEXT,
            created_at TEXT,
            updated_at TEXT,
            UNIQUE(email)
        );
        CREATE TABLE IF NOT EXISTS hotel_contact_list (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            email TEXT NOT NULL,
            hotel_name TEXT,
            created_at TEXT,
            updated_at TEXT,
            UNIQUE(email)
        );
        CREATE TABLE IF NOT EXISTS contract_template (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            chain TEXT NOT NULL,
            template_name TEXT NOT NULL,
            description TEXT,
            filename TEXT,
            file_data BLOB,
            sections TEXT DEFAULT '[]',
            merge_fields TEXT DEFAULT '[]',
            is_active INTEGER DEFAULT 1,
            created_at TEXT DEFAULT (datetime('now')),
            updated_at TEXT DEFAULT (datetime('now'))
        );
        CREATE TABLE IF NOT EXISTS rfp (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            rfp_code TEXT,
            rfp_name TEXT,
            client_org TEXT NOT NULL,
            event_name TEXT,
            start_date TEXT,
            end_date TEXT,
            alt_start_date TEXT,
            alt_end_date TEXT,
            peak_rooms INTEGER,
            total_room_nights INTEGER,
            f_and_b_budget REAL,
            total_attendees INTEGER,
            response_due_date TEXT,
            decision_due_date TEXT,
            status TEXT DEFAULT 'sourcing',
            booking_id TEXT,
            notes TEXT,
            rfp_filename TEXT,
            rfp_data BLOB,
            contract_filename TEXT,
            contract_data BLOB,
            archived INTEGER DEFAULT 0,
            created_at TEXT DEFAULT (datetime('now')),
            updated_at TEXT DEFAULT (datetime('now'))
        );
        CREATE TABLE IF NOT EXISTS rfp_hotel (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            rfp_id INTEGER NOT NULL REFERENCES rfp(id) ON DELETE CASCADE,
            hotel_name TEXT NOT NULL,
            brand TEXT, city TEXT, state TEXT,
            contact_name TEXT, contact_email TEXT,
            contact_phone TEXT, contact_title TEXT,
            status TEXT DEFAULT 'pending',
            proposed_rate REAL,
            commission_pct REAL,
            f_and_b_minimum REAL,
            meeting_room_rental REAL,
            attrition_pct REAL,
            cutoff_days INTEGER,
            concessions TEXT,
            proposal_filename TEXT,
            proposal_data BLOB,
            notes TEXT,
            created_at TEXT DEFAULT (datetime('now')),
            updated_at TEXT DEFAULT (datetime('now'))
        );
        CREATE TABLE IF NOT EXISTS rfp_note (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            rfp_id INTEGER NOT NULL REFERENCES rfp(id) ON DELETE CASCADE,
            rfp_hotel_id INTEGER REFERENCES rfp_hotel(id) ON DELETE SET NULL,
            note_date TEXT NOT NULL,
            note_type TEXT DEFAULT 'internal',
            note_text TEXT NOT NULL,
            created_at TEXT DEFAULT (datetime('now'))
        );
    ''')
    db.commit()
    for _col, _typ in [('ota_url', 'TEXT'), ('cc_emails', "TEXT DEFAULT '[]'"),
                        ('force_current', 'INTEGER DEFAULT 0'),
                        ('hotel_contacts', "TEXT DEFAULT '[]'"),
                        ('force_past', 'INTEGER DEFAULT 0')]:
        try:
            db.execute(f'ALTER TABLE pickup_config ADD COLUMN {_col} {_typ}')
            db.commit()
        except Exception:
            pass
    # Add HHR (Housing History Report) columns to pickup_config
    for _col, _typ in [('hhr_filename', 'TEXT'), ('hhr_data', 'BLOB'), ('hhr_stats', 'TEXT')]:
        try:
            db.execute(f'ALTER TABLE pickup_config ADD COLUMN {_col} {_typ}')
            db.commit()
        except Exception:
            pass
    # Add contract storage + block_is_estimated columns to pickup_config
    for _col, _typ in [('contract_filename', 'TEXT'), ('contract_data', 'BLOB'),
                       ('block_is_estimated', 'INTEGER DEFAULT 0')]:
        try:
            db.execute(f'ALTER TABLE pickup_config ADD COLUMN {_col} {_typ}')
            db.commit()
        except Exception:
            pass
    # Add event_start / event_end date columns to pickup_config
    for _col in ('event_start', 'event_end'):
        try:
            db.execute(f'ALTER TABLE pickup_config ADD COLUMN {_col} TEXT')
            db.commit()
        except Exception:
            pass
    # Add rooming_list_required flag to pickup_config
    try:
        db.execute('ALTER TABLE pickup_config ADD COLUMN rooming_list_required INTEGER DEFAULT 0')
        db.commit()
    except Exception:
        pass
    # Add secondary hotel contact columns to pickup_config
    for _col in ('hotel_contact2', 'hotel_contact2_email'):
        try:
            db.execute(f'ALTER TABLE pickup_config ADD COLUMN {_col} TEXT')
            db.commit()
        except Exception:
            pass
    # Add per-week OTA URL to pickup_weekly
    try:
        db.execute('ALTER TABLE pickup_weekly ADD COLUMN ota_url TEXT')
        db.commit()
    except Exception:
        pass
    # Add CRF columns and alt date columns to rfp/rfp_hotel if not present
    for _tbl, _col, _typ in [
        ('rfp',       'crf_filename',    'TEXT'),
        ('rfp',       'crf_data',        'BLOB'),
        ('rfp_hotel', 'crf_row_data',    'TEXT'),
        ('rfp_hotel', 'crf_version',     'INTEGER DEFAULT 0'),
        ('rfp',       'alt_start_date_2','TEXT'),
        ('rfp',       'alt_end_date_2',  'TEXT'),
        ('rfp',       'alt_start_date_3','TEXT'),
        ('rfp',       'alt_end_date_3',  'TEXT'),
    ]:
        try:
            db.execute(f'ALTER TABLE {_tbl} ADD COLUMN {_col} {_typ}')
            db.commit()
        except Exception:
            pass
    # Add BookingName to ReportPipeline if not present
    try:
        db.execute('ALTER TABLE ReportPipeline ADD COLUMN BookingName TEXT')
        db.commit()
    except Exception:
        pass
    # Add EventID to ReportPipeline if not present
    try:
        db.execute('ALTER TABLE ReportPipeline ADD COLUMN EventID TEXT')
        db.commit()
    except Exception:
        pass


try:
    with app.app_context():
        ensure_pickup_tables()
except Exception:
    pass


# ── Auth tables & seeding ─────────────────────────────────────────────────────

ALL_PERMISSIONS = [
    ('dashboard',                  'Dashboard'),
    ('admin_panel',                'Admin Panel'),
    ('import_bookings',            'Import → Bookings'),
    ('import_payments',            'Import → Payments'),
    ('import_hhr',                 'Import → HHR'),
    ('import_cancelled',           'Import → Cancelled Meetings'),
    ('reports_commission_kristin', 'Missing Pickup/Comm Report (Kristin)'),
    ('reports_commission_team',    'Missing Pickup/Comm Report (Team)'),
    ('reports_payments',           'Reports → Payment Report'),
    ('reports_customer_summary',   'Reports → Customer Summary'),
    ('bookings_view',              'View Bookings'),
    ('bookings_edit',              'Add / Edit Bookings'),
    ('pickups_payments',           'Add Pickups / Payments'),
    ('contracts',                  'Contracts'),
]

_SEED_USERS = [
    ('Peter Wann',    'peter.wann@conferencedirect.com',     'peter',    'CPAinc2026!', 'admin'),
    ('Kristin House', 'kristin.house@conferencedirect.com',  'kristin',  'CPAinc2026!', 'admin'),
    ('Geralyn Krist', 'geralyn.krist@conferencedirect.com',  'geralyn',  'CPAinc2026!', 'user'),
    ('Morgan Basham', 'morgan.basham@conferencedirect.com',  'morgan',   'CPAinc2026!', 'user'),
    ('Ashleigh Buhr', 'ashleigh.buhr@conferencedirect.com', 'ashleigh', 'CPAinc2026!', 'user'),
]


def ensure_auth_tables():
    db = get_db()
    db.executescript('''
        CREATE TABLE IF NOT EXISTS Users (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            name          TEXT    NOT NULL,
            email         TEXT    NOT NULL UNIQUE,
            username      TEXT    NOT NULL UNIQUE,
            password_hash TEXT    NOT NULL,
            role          TEXT    NOT NULL DEFAULT 'user',
            active        INTEGER NOT NULL DEFAULT 1
        );
        CREATE TABLE IF NOT EXISTS UserPermissions (
            id         INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id    INTEGER NOT NULL REFERENCES Users(id) ON DELETE CASCADE,
            permission TEXT    NOT NULL,
            enabled    INTEGER NOT NULL DEFAULT 0,
            UNIQUE(user_id, permission)
        );
        CREATE TABLE IF NOT EXISTS UserAccountAccess (
            id           INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id      INTEGER NOT NULL REFERENCES Users(id) ON DELETE CASCADE,
            account_name TEXT    NOT NULL,
            UNIQUE(user_id, account_name)
        );
        CREATE TABLE IF NOT EXISTS UserMicrosoftTokens (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id       INTEGER NOT NULL UNIQUE REFERENCES Users(id) ON DELETE CASCADE,
            access_token  TEXT    NOT NULL,
            refresh_token TEXT    NOT NULL,
            expires_at    REAL    NOT NULL,
            ms_user_email TEXT,
            connected_at  TEXT    DEFAULT (datetime('now'))
        );
        CREATE TABLE IF NOT EXISTS UserPipelineAssociates (
            id             INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id        INTEGER NOT NULL REFERENCES Users(id) ON DELETE CASCADE,
            associate_name TEXT    NOT NULL,
            UNIQUE(user_id, associate_name)
        );
        CREATE TABLE IF NOT EXISTS status_board_ignore (
            id           INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id      INTEGER NOT NULL REFERENCES Users(id) ON DELETE CASCADE,
            config_id    INTEGER NOT NULL REFERENCES pickup_config(id) ON DELETE CASCADE,
            issue_type   TEXT NOT NULL,
            ignore_until TEXT,
            created_at   TEXT DEFAULT (datetime('now')),
            UNIQUE(user_id, config_id, issue_type)
        );
        CREATE TABLE IF NOT EXISTS activity_log (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            username    TEXT,
            user_id     INTEGER,
            timestamp   TEXT DEFAULT (datetime('now','localtime')),
            method      TEXT,
            endpoint    TEXT,
            path        TEXT,
            description TEXT,
            detail      TEXT,
            ip_address  TEXT
        );
        CREATE TABLE IF NOT EXISTS pickup_tasks (
            id                  INTEGER PRIMARY KEY AUTOINCREMENT,
            config_id           INTEGER REFERENCES pickup_config(id) ON DELETE CASCADE,
            title               TEXT NOT NULL,
            due_date            TEXT,
            assigned_to_user_id INTEGER REFERENCES Users(id),
            assigned_to_name    TEXT,
            created_by_user_id  INTEGER,
            created_by_name     TEXT,
            created_at          TEXT DEFAULT (datetime('now','localtime')),
            completed_at        TEXT,
            completed_by        TEXT,
            task_notes          TEXT
        );
        CREATE TABLE IF NOT EXISTS pickup_event_notes (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            config_id   INTEGER REFERENCES pickup_config(id) ON DELETE CASCADE,
            user_id     INTEGER,
            username    TEXT,
            note_text   TEXT NOT NULL,
            created_at  TEXT DEFAULT (datetime('now','localtime'))
        );
        CREATE TABLE IF NOT EXISTS pickup_weekly_deleted (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            original_id     INTEGER,
            config_id       INTEGER,
            report_date     TEXT,
            pickup_by_night TEXT,
            total_rooms     INTEGER,
            change_from_last INTEGER,
            pct_of_block    REAL,
            pct_of_attrition REAL,
            ota_rate        REAL,
            label           TEXT,
            notes           TEXT,
            deleted_at      TEXT DEFAULT (datetime('now','localtime')),
            deleted_by      TEXT
        );
        CREATE TABLE IF NOT EXISTS pickup_change_log (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            config_id   INTEGER REFERENCES pickup_config(id) ON DELETE CASCADE,
            user_id     INTEGER,
            username    TEXT,
            timestamp   TEXT DEFAULT (datetime('now','localtime')),
            action      TEXT,
            field       TEXT,
            old_value   TEXT,
            new_value   TEXT
        );
    ''')
    db.commit()
    _seed_users(db)
    _seed_pipeline_associates(db)


def _seed_pipeline_associates(db):
    """Ensure Peter and Kristin are mapped to Kristin House + Morgan Klinkradt associates."""
    for username in ('peter', 'kristin'):
        row = db.execute('SELECT id FROM Users WHERE username=?', (username,)).fetchone()
        if not row:
            continue
        uid = row['id']
        for assoc in ('Kristin House', 'Morgan Klinkradt'):
            db.execute(
                'INSERT OR IGNORE INTO UserPipelineAssociates (user_id, associate_name) VALUES (?,?)',
                (uid, assoc)
            )

    # Remove pickup_config records that don't belong in this system
    # (Peter Wann's FCCS accounts and any non-Kristin/Morgan orgs entered in error)
    non_system_orgs = ('FCCS',)
    for org in non_system_orgs:
        db.execute(
            "DELETE FROM pickup_config WHERE organization=? "
            "AND id NOT IN (SELECT config_id FROM pickup_weekly) "
            "AND id NOT IN (SELECT config_id FROM pickup_contact_log)",
            (org,)
        )

    db.commit()


def _hash_password(password):
    return generate_password_hash(password, method='pbkdf2:sha256')


def _seed_users(db):
    for name, email, username, password, role in _SEED_USERS:
        existing = db.execute('SELECT id FROM Users WHERE email = ? OR username = ?', (email, username)).fetchone()
        if not existing:
            ph = _hash_password(password)
            try:
                db.execute(
                    'INSERT INTO Users (name, email, username, password_hash, role) VALUES (?,?,?,?,?)',
                    (name, email, username, ph, role)
                )
            except Exception:
                continue
            db.commit()
            user = db.execute('SELECT id FROM Users WHERE email = ?', (email,)).fetchone()
            uid = user['id']
            is_admin = (role == 'admin')
            for perm, _ in ALL_PERMISSIONS:
                enabled = 1 if is_admin or perm in ('bookings_view', 'pickups_payments') else 0
                db.execute(
                    'INSERT OR IGNORE INTO UserPermissions (user_id, permission, enabled) VALUES (?,?,?)',
                    (uid, perm, enabled)
                )
            db.commit()


try:
    with app.app_context():
        ensure_auth_tables()
except Exception:
    pass


# ── Auth helpers ──────────────────────────────────────────────────────────────

def get_current_user():
    uid = session.get('user_id')
    if not uid:
        return None
    return get_db().execute('SELECT * FROM Users WHERE id = ? AND active = 1', (uid,)).fetchone()


def has_permission(user, perm):
    if user is None:
        return False
    if user['role'] == 'admin':
        return True
    row = get_db().execute(
        'SELECT enabled FROM UserPermissions WHERE user_id = ? AND permission = ?',
        (user['id'], perm)
    ).fetchone()
    return bool(row and row['enabled'])


def get_user_account_filter(user=None):
    """Returns a list of account names the user may see, or None for admins (no filter).
    Used for financial reporting — admins always see all."""
    if user is None:
        user = get_current_user()
    if user is None:
        return []
    if user['role'] == 'admin':
        return None
    rows = get_db().execute(
        'SELECT account_name FROM UserAccountAccess WHERE user_id = ?', (user['id'],)
    ).fetchall()
    return [r['account_name'] for r in rows]


def get_pickup_account_filter(user=None, db=None):
    """Returns an account-name filter list for Pickup, RFP, and Status Board views.
    Unlike get_user_account_filter, this applies to ALL users including admins.

    Priority:
      1. UserPipelineAssociates — dynamically queries ReportPipeline for all
         AccountNames belonging to those associates. Used for Peter & Kristin.
      2. UserAccountAccess — explicit list (Morgan, Ashleigh, Geralyn).
      3. None — no filter (see everything). Only if user has neither table populated.
    """
    if user is None:
        user = get_current_user()
    if user is None:
        return []
    if db is None:
        db = get_db()

    # 1. Associate-driven filter (pipeline lookup)
    assoc_rows = db.execute(
        'SELECT associate_name FROM UserPipelineAssociates WHERE user_id=?', (user['id'],)
    ).fetchall()
    if assoc_rows:
        associates = [r['associate_name'] for r in assoc_rows]
        ph = ','.join('?' * len(associates))
        acct_rows = db.execute(
            f"SELECT DISTINCT AccountName FROM ReportPipeline "
            f"WHERE BookingAssociate IN ({ph}) AND AccountName IS NOT NULL",
            associates
        ).fetchall()
        return [r['AccountName'] for r in acct_rows]

    # 2. Explicit UserAccountAccess list
    acct_rows = db.execute(
        'SELECT account_name FROM UserAccountAccess WHERE user_id=?', (user['id'],)
    ).fetchall()
    if acct_rows:
        return [r['account_name'] for r in acct_rows]

    # 3. No restriction (admin with no filter configured)
    if user['role'] == 'admin':
        return None

    return []  # non-admin with no access configured


@app.before_request
def require_login():
    exempt = {'login', 'logout', 'static'}
    if request.endpoint in exempt or request.endpoint is None:
        return
    uid = session.get('user_id')
    if not uid:
        if request.headers.get('X-Requested-With') == 'fetch' or request.is_json:
            return jsonify({'ok': False, 'error': 'Not authenticated'}), 401
        return redirect(url_for('login', next=request.path))
    # ── Session timeout (1 hour of inactivity) ──
    last_active = session.get('last_activity')
    now_ts = datetime.utcnow()
    if last_active:
        try:
            elapsed = (now_ts - datetime.fromisoformat(last_active)).total_seconds()
            if elapsed > SESSION_TIMEOUT_SECONDS:
                session.clear()
                if request.headers.get('X-Requested-With') == 'fetch' or request.is_json:
                    return jsonify({'ok': False, 'error': 'Session expired'}), 401
                flash('Your session expired after 1 hour of inactivity. Please log in again.', 'warning')
                return redirect(url_for('login', next=request.path))
        except Exception:
            pass
    session['last_activity'] = now_ts.isoformat()
    user = get_db().execute('SELECT id, active FROM Users WHERE id = ?', (uid,)).fetchone()
    if not user or not user['active']:
        session.clear()
        return redirect(url_for('login'))


# ── Change log helper ────────────────────────────────────────────────────────

def _log_change(db, config_id, action, changes):
    """Record field-level changes to a pickup_config record.
    changes = list of (field, old_value, new_value) tuples.
    Only logs entries where old != new."""
    uid      = session.get('user_id')
    username = session.get('username') or 'unknown'
    for field, old_val, new_val in changes:
        old_s = str(old_val) if old_val is not None else ''
        new_s = str(new_val) if new_val is not None else ''
        if old_s != new_s:
            db.execute(
                '''INSERT INTO pickup_change_log (config_id, user_id, username, action, field, old_value, new_value)
                   VALUES (?,?,?,?,?,?,?)''',
                (config_id, uid, username, action, field, old_s or None, new_s or None)
            )


# ── Activity logging ──────────────────────────────────────────────────────────

# Human-readable labels for each endpoint (keyed by actual Flask function name)
_ENDPOINT_LABELS = {
    # ── Auth ──
    'login':                            'Logged in',
    'logout':                           'Logged out',
    # ── Home / Bookings ──
    'index':                            'Viewed Home / Dashboard',
    'bookings':                         'Viewed Bookings list',
    'booking_detail':                   'Viewed Booking #{booking_id}',
    'booking_new':                      'Created new Booking',
    'booking_edit':                     'Edited Booking #{booking_id}',
    'booking_cancel':                   'Cancelled Booking #{booking_id}',
    # ── Contracts ──
    'booking_contract_upload':          'Uploaded contract — Booking #{booking_id}',
    'booking_contract_download':        'Downloaded contract — Booking #{booking_id}',
    'booking_contract_delete':          'Deleted contract — Booking #{booking_id}',
    'booking_contract_extract':         'Extracted contract to Pickup — Booking #{booking_id}',
    'booking_contract_select_pickup':   'Selected Pickup for contract — Booking #{booking_id}',
    'booking_contract_parse_for_pickup':'Parsed contract into Pickup (config #{cid})',
    # ── Imports ──
    'import_bookings':                  'Imported bookings from Excel',
    'import_cancelled':                 'Imported cancelled bookings',
    'import_payments':                  'Imported payments',
    'import_voucher':                   'Imported voucher data',
    'import_hhr':                       'Imported Housing History Report (HHR)',
    # ── Pickup Dashboard / Events ──
    'pickup_dashboard':                 'Viewed Pickup Tracking dashboard',
    'pickup_new_event':                 'Created new Pickup event',
    'pickup_event':                     'Viewed Pickup event (config #{cid})',
    'pickup_edit_event':                'Edited Pickup event (config #{cid})',
    'pickup_sync_pipeline':             'Synced Pickup from Pipeline',
    'pickup_fill_missing':              'Filled missing Pickup data',
    'pickup_auto_match':                'Auto-matched Pickup records',
    'pickup_import_xlsx':               'Imported Pickup data from Excel',
    'pickup_import_xlsx_confirm':       'Confirmed Pickup Excel import',
    # ── Pickup Reports ──
    'pickup_weekly_new':                'Added pickup report (config #{cid})',
    'pickup_weekly_edit':               'Edited pickup report (config #{cid})',
    'pickup_weekly_delete':             'Deleted pickup report (config #{cid})',
    'pickup_event_report':              'Viewed Pickup report (config #{primary_cid})',
    'pickup_event_report_xlsx':         'Downloaded Pickup report Excel (config #{primary_cid})',
    # ── Pickup Contracts / HHR ──
    'pickup_upload_contract':           'Uploaded contract — Pickup config #{cid}',
    'pickup_contract_download':         'Downloaded contract — Pickup config #{cid}',
    'pickup_hhr_download':              'Downloaded HHR — Pickup config #{cid}',
    'pickup_housing_form':              'Viewed Housing Form (config #{cid})',
    'pickup_final_history':             'Viewed Final History (config #{cid})',
    # ── Pickup Contacts / Emails ──
    'pickup_contact_log':               'Logged hotel contact (config #{cid})',
    'pickup_email_housing':             'Sent housing email (config #{cid})',
    'pickup_email_hotel':               'Sent hotel email (config #{cid})',
    'pickup_email_client':              'Sent client email (config #{cid})',
    'pickup_email_client_launch_outlook':'Launched Outlook for client email (config #{cid})',
    'pickup_email_client_paste':        'Copied client email to clipboard (config #{cid})',
    'pickup_email_post_report':         'Sent post-event report (config #{cid})',
    # ── Pickup Rooming List ──
    'pickup_rooming_upload':            'Uploaded Rooming List (config #{cid})',
    'pickup_rooming_manual':            'Added manual Rooming List entry (config #{cid})',
    'pickup_rooming_review':            'Reviewed Rooming List (config #{cid})',
    'pickup_rooming_confirm':           'Confirmed Rooming List (config #{cid})',
    'pickup_rooming_download_csv':      'Downloaded Rooming List CSV (config #{cid})',
    # ── Customer Reports ──
    'pickup_customer_report_select':    'Selected accounts for Customer Report',
    'pickup_customer_report_xlsx':      'Downloaded Customer Report Excel',
    'pickup_export_row':                'Exported pickup row (config #{cid})',
    # ── Status Board ──
    'status_board':                     'Viewed Status Board',
    'status_board_ignore':              'Ignored Status Board item',
    # ── Admin ──
    'admin_users':                      'Viewed Admin — User Management',
    'admin_toggle_permission':          'Changed permission for user #{uid}',
    'admin_update_accounts':            'Updated account access for user #{uid}',
    'admin_toggle_active':              'Toggled active status for user #{uid}',
    'admin_reset_password':             'Reset password for user #{uid}',
    'admin_download_db':                'Downloaded full database backup',
    'admin_export_pickup_data':         'Exported pickup data (JSON)',
    'admin_import_pickup_data':         'Imported pickup data (JSON)',
    'admin_activity_log':               'Viewed Activity Log',
    # ── RFP ──
    'rfp_list':                         'Viewed RFP list',
    'rfp_detail':                       'Viewed RFP #{rfp_id}',
    'rfp_create':                       'Created new RFP',
    'rfp_edit':                         'Edited RFP #{rfp_id}',
    'rfp_delete':                       'Deleted RFP #{rfp_id}',
}

# Endpoints to skip logging (too noisy or not meaningful)
_SKIP_LOG_ENDPOINTS = {
    'static', None,
    'admin_activity_log',   # avoid self-referential noise
    'status_board_ignore',  # logged via POST action description instead
}

def _log_activity(response):
    """Record the current request to activity_log. Called from after_request."""
    try:
        endpoint = request.endpoint
        if endpoint in _SKIP_LOG_ENDPOINTS:
            return
        # Only log successful page loads and meaningful POSTs (skip redirects for GET)
        if request.method == 'GET' and response.status_code in (301, 302):
            return
        # Skip non-HTML/non-download responses that are just JSON fragments
        ct = response.content_type or ''
        if 'application/json' in ct and request.method == 'GET':
            return

        uid      = session.get('user_id')
        username = session.get('username')
        if not username and uid:
            # Fall back to DB lookup for sessions that predate the username cache
            try:
                row = get_db().execute('SELECT name FROM Users WHERE id=?', (uid,)).fetchone()
                if row:
                    username = row['name']
                    session['username'] = username  # cache it going forward
            except Exception:
                pass
        username = username or 'unknown'
        path     = request.path

        # Build human-readable description
        label = _ENDPOINT_LABELS.get(endpoint)
        if label:
            # Fill in URL params
            kw = dict(request.view_args or {})
            try:
                description = label.format(**kw)
            except KeyError:
                description = label
        else:
            # Fallback: turn endpoint name into readable text
            description = (endpoint or path).replace('_', ' ').title()
            if request.method == 'POST':
                description = 'Submitted: ' + description

        # Extra detail: form keys (not values) for POSTs, or query string for GETs
        detail = None
        if request.method == 'POST':
            keys = [k for k in (request.form.keys() if request.form else [])
                    if k not in ('password', 'password_hash', 'csrf_token')]
            if keys:
                detail = 'Fields: ' + ', '.join(keys[:8])
        elif request.args:
            detail = 'Query: ' + '&'.join(f'{k}={v}' for k, v in list(request.args.items())[:5])

        ip = request.headers.get('X-Forwarded-For', request.remote_addr or '')
        if ip and ',' in ip:
            ip = ip.split(',')[0].strip()

        db = get_db()
        db.execute(
            '''INSERT INTO activity_log (username, user_id, method, endpoint, path, description, detail, ip_address)
               VALUES (?,?,?,?,?,?,?,?)''',
            (username, uid, request.method, endpoint, path, description, detail, ip)
        )
        db.commit()
    except Exception:
        pass  # Never let logging break the actual response


@app.after_request
def after_request_log(response):
    _log_activity(response)
    return response


@app.context_processor
def inject_user():
    user = get_current_user()
    def _has_perm(perm):
        return has_permission(user, perm)
    return dict(current_user=user, has_perm=_has_perm, all_permissions=ALL_PERMISSIONS)


# ── Login / Logout ────────────────────────────────────────────────────────────

@app.route('/login', methods=['GET', 'POST'])
def login():
    if session.get('user_id'):
        return redirect(url_for('index'))
    error = None
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '')
        db = get_db()
        user = db.execute(
            'SELECT * FROM Users WHERE (username = ? OR email = ?) AND active = 1',
            (username, username)
        ).fetchone()
        if user and check_password_hash(user['password_hash'], password):
            session.clear()
            session['user_id']   = user['id']
            session['username']  = user['name']
            session['user_role'] = user['role']
            next_page = request.form.get('next') or request.args.get('next') or url_for('status_board')
            return redirect(next_page)
        error = 'Invalid username or password.'
    return render_template('login.html', error=error, next=request.args.get('next', ''))


@app.route('/help')
def help_page():
    return render_template('help.html')

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))


# ── Admin User Management ─────────────────────────────────────────────────────

@app.route('/admin/users')
def admin_users():
    user = get_current_user()
    if not has_permission(user, 'admin_panel'):
        flash('Access denied.', 'error')
        return redirect(url_for('index'))
    db = get_db()
    users = db.execute('SELECT * FROM Users ORDER BY role DESC, name').fetchall()
    all_perms_rows = db.execute('SELECT user_id, permission, enabled FROM UserPermissions').fetchall()
    perm_map = {}
    for r in all_perms_rows:
        perm_map.setdefault(r['user_id'], {})[r['permission']] = bool(r['enabled'])
    available_accounts = [r[0] for r in db.execute(
        "SELECT DISTINCT AccountName FROM ReportPipeline "
        "WHERE AccountName IS NOT NULL "
        "AND LOWER(BookingAssociate) = 'kristin house' "
        "ORDER BY AccountName"
    ).fetchall()]
    acc_rows = db.execute('SELECT user_id, account_name FROM UserAccountAccess').fetchall()
    acc_map = {}
    for r in acc_rows:
        acc_map.setdefault(r['user_id'], [])
        acc_map[r['user_id']].append(r['account_name'])
    return render_template('admin_users.html',
        users=users, perm_map=perm_map,
        all_permissions=ALL_PERMISSIONS,
        available_accounts=available_accounts,
        acc_map=acc_map)


@app.route('/admin/users/<int:uid>/permission/<perm>/toggle', methods=['POST'])
def admin_toggle_permission(uid, perm):
    user = get_current_user()
    if not has_permission(user, 'admin_panel'):
        return jsonify({'ok': False, 'error': 'Access denied'}), 403
    db = get_db()
    existing = db.execute(
        'SELECT enabled FROM UserPermissions WHERE user_id = ? AND permission = ?', (uid, perm)
    ).fetchone()
    if existing:
        new_val = 0 if existing['enabled'] else 1
        db.execute('UPDATE UserPermissions SET enabled = ? WHERE user_id = ? AND permission = ?',
                   (new_val, uid, perm))
    else:
        new_val = 1
        db.execute('INSERT INTO UserPermissions (user_id, permission, enabled) VALUES (?,?,?)',
                   (uid, perm, new_val))
    db.commit()
    return jsonify({'ok': True, 'enabled': bool(new_val)})


@app.route('/admin/users/<int:uid>/accounts', methods=['POST'])
def admin_update_accounts(uid):
    user = get_current_user()
    if not has_permission(user, 'admin_panel'):
        return jsonify({'ok': False, 'error': 'Access denied'}), 403
    db = get_db()
    data = request.get_json(silent=True) or {}
    accounts = data.get('accounts', [])
    db.execute('DELETE FROM UserAccountAccess WHERE user_id = ?', (uid,))
    for acct in accounts:
        if acct.strip():
            db.execute('INSERT OR IGNORE INTO UserAccountAccess (user_id, account_name) VALUES (?,?)',
                       (uid, acct.strip()))
    db.commit()
    return jsonify({'ok': True, 'count': len(accounts)})


@app.route('/admin/users/<int:uid>/active/toggle', methods=['POST'])
def admin_toggle_active(uid):
    user = get_current_user()
    if not has_permission(user, 'admin_panel'):
        return jsonify({'ok': False}), 403
    db = get_db()
    existing = db.execute('SELECT active FROM Users WHERE id = ?', (uid,)).fetchone()
    if not existing:
        return jsonify({'ok': False}), 404
    new_val = 0 if existing['active'] else 1
    db.execute('UPDATE Users SET active = ? WHERE id = ?', (new_val, uid))
    db.commit()
    return jsonify({'ok': True, 'active': bool(new_val)})


@app.route('/admin/users/<int:uid>/password', methods=['POST'])
def admin_reset_password(uid):
    user = get_current_user()
    if not has_permission(user, 'admin_panel'):
        return jsonify({'ok': False}), 403
    data = request.get_json(silent=True) or {}
    new_password = data.get('password', '')
    if len(new_password) < 6:
        return jsonify({'ok': False, 'error': 'Password must be at least 6 characters.'}), 400
    ph = _hash_password(new_password)
    db = get_db()
    db.execute('UPDATE Users SET password_hash = ? WHERE id = ?', (ph, uid))
    db.commit()
    return jsonify({'ok': True})


@app.route('/admin/download-db')
def admin_download_db():
    """Stream a consistent snapshot of the live SQLite database (admin only)."""
    user = get_current_user()
    if not user or not has_permission(user, 'admin_panel'):
        return 'Forbidden — admin login required.', 403
    import sqlite3 as _sq3, tempfile, os as _os
    from flask import Response
    with tempfile.NamedTemporaryFile(suffix='.sqlite', delete=False) as tf:
        tmp_path = tf.name
    try:
        src = _sq3.connect(DATABASE)
        dst = _sq3.connect(tmp_path)
        src.backup(dst)          # atomic consistent copy
        src.close()
        dst.close()
        with open(tmp_path, 'rb') as f:
            db_bytes = f.read()
    finally:
        try:
            _os.unlink(tmp_path)
        except Exception:
            pass
    return Response(
        db_bytes,
        mimetype='application/octet-stream',
        headers={'Content-Disposition': 'attachment; filename="CPAinc.sqlite"'}
    )


@app.route('/admin/activity-log')
def admin_activity_log():
    user = get_current_user()
    if not user or user['role'] != 'admin':
        flash('Admin access required.', 'error')
        return redirect(url_for('index'))

    db = get_db()

    # Filters
    filter_user = request.args.get('user', '').strip()
    filter_date = request.args.get('date', '').strip()
    filter_endpoint = request.args.get('endpoint', '').strip()
    page = max(1, int(request.args.get('page', 1)))
    per_page = 100

    where  = []
    params = []
    if filter_user:
        where.append('username = ?')
        params.append(filter_user)
    if filter_date:
        where.append("DATE(timestamp) = ?")
        params.append(filter_date)
    if filter_endpoint:
        where.append('endpoint = ?')
        params.append(filter_endpoint)

    where_sql = ('WHERE ' + ' AND '.join(where)) if where else ''

    total = db.execute(
        f'SELECT COUNT(*) FROM activity_log {where_sql}', params
    ).fetchone()[0]

    logs = db.execute(
        f'''SELECT id, username, timestamp, method, endpoint, path, description, detail, ip_address
            FROM activity_log {where_sql}
            ORDER BY id DESC
            LIMIT ? OFFSET ?''',
        params + [per_page, (page - 1) * per_page]
    ).fetchall()

    # Distinct users and endpoints for filter dropdowns
    all_users     = [r[0] for r in db.execute(
        "SELECT DISTINCT username FROM activity_log WHERE username IS NOT NULL ORDER BY username"
    ).fetchall()]
    all_endpoints = [r[0] for r in db.execute(
        "SELECT DISTINCT endpoint FROM activity_log WHERE endpoint IS NOT NULL ORDER BY endpoint"
    ).fetchall()]

    total_pages = max(1, (total + per_page - 1) // per_page)

    return render_template('activity_log.html',
        logs=logs,
        total=total,
        page=page,
        total_pages=total_pages,
        per_page=per_page,
        filter_user=filter_user,
        filter_date=filter_date,
        filter_endpoint=filter_endpoint,
        all_users=all_users,
        all_endpoints=all_endpoints,
    )


@app.route('/admin/export-pickup-data')
def admin_export_pickup_data():
    """Export all pickup_weekly rows as JSON keyed by booking_id (safe to share)."""
    user = get_current_user()
    if not user or not has_permission(user, 'admin_panel'):
        return 'Forbidden — admin login required.', 403
    db = get_db()
    rows = db.execute('''
        SELECT pc.booking_id, pw.report_date, pw.pickup_by_night, pw.total_rooms,
               pw.pct_of_block, pw.pct_of_attrition, pw.ota_rate, pw.label, pw.notes
        FROM pickup_weekly pw
        JOIN pickup_config pc ON pc.id = pw.config_id
        ORDER BY pc.booking_id, pw.report_date
    ''').fetchall()
    out = [dict(r) for r in rows]
    resp = app.response_class(
        response=json.dumps(out, indent=2),
        mimetype='application/json',
        headers={'Content-Disposition': 'attachment; filename="pickup_weekly_export.json"'}
    )
    return resp


@app.route('/admin/import-pickup-data', methods=['GET', 'POST'])
def admin_import_pickup_data():
    """Import pickup_weekly rows from JSON — matches by booking_id, never overwrites existing rows."""
    user = get_current_user()
    if not user or not has_permission(user, 'admin_panel'):
        return 'Forbidden — admin login required.', 403
    if request.method == 'GET':
        return '''<!doctype html><html><body style="font-family:sans-serif;padding:2rem">
        <h3>Import Pickup Weekly Data</h3>
        <p>Uploads pickup_weekly rows matched by booking_id. Skips duplicates (same booking_id + report_date).</p>
        <form method="post" enctype="multipart/form-data">
          <input type="file" name="json_file" accept=".json"><br><br>
          <button type="submit">Import</button>
        </form></body></html>'''
    # POST — process uploaded JSON file
    f = request.files.get('json_file')
    if not f:
        return 'No file provided.', 400
    try:
        data = json.loads(f.read().decode('utf-8'))
    except Exception as e:
        return f'Invalid JSON: {e}', 400
    db = get_db()
    # Build booking_id → config_id map from THIS database
    cfg_map = {}
    for row in db.execute('SELECT id, booking_id FROM pickup_config'):
        if row['booking_id']:
            cfg_map[str(row['booking_id'])] = row['id']
    # Get existing (config_id, report_date) pairs to skip duplicates
    existing = set()
    for row in db.execute('SELECT config_id, report_date FROM pickup_weekly'):
        existing.add((row['config_id'], row['report_date']))
    inserted = skipped_no_config = skipped_duplicate = 0
    for row in data:
        bid = str(row.get('booking_id', ''))
        cid = cfg_map.get(bid)
        if not cid:
            skipped_no_config += 1
            continue
        rd = row.get('report_date', '')
        if (cid, rd) in existing:
            skipped_duplicate += 1
            continue
        db.execute('''INSERT INTO pickup_weekly
            (config_id, report_date, pickup_by_night, total_rooms,
             pct_of_block, pct_of_attrition, ota_rate, label, notes)
            VALUES (?,?,?,?,?,?,?,?,?)''',
            (cid, rd, row.get('pickup_by_night'), row.get('total_rooms'),
             row.get('pct_of_block'), row.get('pct_of_attrition'),
             row.get('ota_rate'), row.get('label'), row.get('notes')))
        existing.add((cid, rd))
        inserted += 1
    db.commit()
    return (f'<p>Done. Inserted: <strong>{inserted}</strong> rows. '
            f'Skipped duplicates: {skipped_duplicate}. '
            f'Skipped (event not found here): {skipped_no_config}.</p>'
            f'<a href="/pickup">Go to Pickup Tracking</a>')


@app.route('/admin/fix-asa-cssa-sssa-dates')
def admin_fix_asa_dates():
    """One-time fix: shift contracted_block dates and correct booking_ids for ASA-CSSA-SSSA 2027/2028 records."""
    user = get_current_user()
    if not user or not has_permission(user, 'admin_panel'):
        return 'Forbidden — admin login required.', 403

    db = get_db()
    results = []

    # Correct booking_id mapping by hotel name substring (case-insensitive)
    HOTEL_FIXES_2027 = [
        ('hilton columbus',           '185599', 'Hilton Columbus Downtown'),
        ('hyatt regency columbus',    '185603', 'Hyatt Regency Columbus'),
        ('drury',                     '185601', 'Drury Inn & Suites Convention Centre'),
        ('hampton inn',               '185524', 'Hampton Inn & Suites Columbus-Downtown'),
        ('ac hotel columbus',         '185521', 'AC Hotel Columbus Downtown'),
        ('ac columbus',               '185521', 'AC Hotel Columbus Downtown'),
        ('red roof',                  '185879', 'Red Roof PLUS+ Columbus Downtown'),
    ]
    HOTEL_FIXES_2028 = [
        ('omni',                      '193163', 'Omni Fort Lauderdale'),
        ('hilton marina',             '193346', 'Hilton Fort Lauderdale Marina'),
        ('hilton fort lauderdale',    '193346', 'Hilton Fort Lauderdale Marina'),
        ('renaissance',               '75067',  'Renaissance Fort Lauderdale'),
        ('embassy suites',            '193345', 'Embassy Suites by Hilton Fort Lauderdale'),
        ('marriott harbor',           '192896', 'Fort Lauderdale Marriott Harbor Beach'),
        ('hyatt place',               '193347', 'Hyatt Place Fort Lauderdale Airport'),
    ]

    def shift_block_year(block_json, old_year, new_year):
        """Replace old_year with new_year in all date keys of a contracted_block JSON string."""
        try:
            block = json.loads(block_json or '{}')
        except Exception:
            block = {}
        new_block = {}
        for k, v in block.items():
            new_k = k.replace(str(old_year) + '-', str(new_year) + '-', 1)
            new_block[new_k] = v
        return json.dumps(new_block)

    def fix_group(year_label, old_year, new_year, hotel_fixes, event_new_name, org_name):
        rows = db.execute(
            "SELECT id, hotel, contracted_block, booking_id FROM pickup_config "
            "WHERE LOWER(event_name) LIKE ? AND status != 'archived'",
            (f'%asa%cssa%sssa%{year_label}%',)
        ).fetchall()
        for row in rows:
            hotel_lower = (row['hotel'] or '').lower()
            new_bid = None
            new_hotel_name = row['hotel']
            for key, bid, clean_name in hotel_fixes:
                if key in hotel_lower:
                    new_bid = bid
                    new_hotel_name = clean_name
                    break
            new_block = shift_block_year(row['contracted_block'], old_year, new_year)
            db.execute(
                '''UPDATE pickup_config
                   SET booking_id=COALESCE(?,booking_id),
                       hotel=?,
                       event_name=?,
                       organization=?,
                       contracted_block=?,
                       status='active'
                   WHERE id=?''',
                (new_bid, new_hotel_name, event_new_name, org_name, new_block, row['id'])
            )
            results.append(f"ID {row['id']}: {row['hotel']} → {new_hotel_name} | "
                           f"booking_id={new_bid or row['booking_id']} | "
                           f"block shifted {old_year}→{new_year}")

    fix_group('2027', 2022, 2027, HOTEL_FIXES_2027,
              '2027 ASA-CSSA-SSSA Annual Meeting',
              'Alliance of Crop, Soil and Environmental Science Societies')
    fix_group('2028', 2023, 2028, HOTEL_FIXES_2028,
              '2028 ASA-CSSA-SSSA Annual Meeting',
              'Alliance of Crop, Soil and Environmental Science Societies')

    # Also fix any remaining record with 2022 block dates under any ASA CSSA SSSA event
    leftover = db.execute(
        "SELECT id, hotel, contracted_block, event_name FROM pickup_config "
        "WHERE LOWER(event_name) LIKE '%asa%cssa%sssa%2022%'"
    ).fetchall()
    for row in leftover:
        new_block = shift_block_year(row['contracted_block'], 2022, 2027)
        db.execute(
            "UPDATE pickup_config SET contracted_block=?, event_name=?, status='active' WHERE id=?",
            (new_block, '2027 ASA-CSSA-SSSA Annual Meeting', row['id'])
        )
        results.append(f"ID {row['id']}: leftover 2022 event fixed → 2027 block | hotel={row['hotel']}")

    db.commit()

    html = '<h3>ASA-CSSA-SSSA Date Fix — Results</h3>'
    if results:
        html += '<ul>' + ''.join(f'<li>{r}</li>' for r in results) + '</ul>'
        html += f'<p><strong>{len(results)} record(s) updated.</strong></p>'
    else:
        html += '<p>No matching records found — nothing to fix (or already fixed).</p>'
    html += '<p><a href="/pickup">Go to Pickup Tracking</a></p>'
    return html


@app.route('/admin/fix-wrong-year-blocks')
def admin_fix_wrong_year_blocks():
    """One-time fix: clear contracted_block dates that don't match the booking's ReportPipeline year,
    and unarchive the record so it reappears in the status board with correct pipeline-date fallback."""
    user = get_current_user()
    if not user or not has_permission(user, 'admin_panel'):
        return 'Forbidden — admin login required.', 403

    db  = get_db()
    results = []

    # These booking IDs were reported as having date inconsistencies.
    # For each: look up the correct year from ReportPipeline, then clear contracted_block
    # if its date keys belong to a different year.
    TARGET_BIDS = ['170939', '170973', '212752', '231248', '226442', '231993']

    pipeline = {}
    for row in db.execute(
        "SELECT BookingId, StartDate FROM ReportPipeline WHERE BookingId IN (?,?,?,?,?,?)",
        TARGET_BIDS
    ).fetchall():
        s = _to_iso(row['StartDate']) or ''
        if len(s) >= 4:
            pipeline[str(row['BookingId'])] = int(s[:4])  # correct year

    for bid in TARGET_BIDS:
        correct_year = pipeline.get(bid)
        rows = db.execute(
            "SELECT id, event_name, hotel, contracted_block, status FROM pickup_config "
            "WHERE booking_id = ?", (bid,)
        ).fetchall()
        for row in rows:
            cid   = row['id']
            block = {}
            try:
                block = json.loads(row['contracted_block'] or '{}')
            except Exception:
                pass

            stray_keys  = [k for k in block if len(k) >= 4 and correct_year and int(k[:4]) != correct_year]
            good_keys   = [k for k in block if k not in stray_keys]
            need_unarch = row['status'] == 'archived'

            if stray_keys or need_unarch:
                if stray_keys and not good_keys:
                    # All keys are wrong-year → wipe entire block
                    new_block = '{}'
                    block_note = f"cleared entire block (stray dates: {stray_keys})"
                elif stray_keys:
                    # Mixed: remove only the stray keys, keep the correct ones
                    clean = {k: v for k, v in block.items() if k not in stray_keys}
                    new_block = json.dumps(clean)
                    block_note = f"removed stray date(s): {stray_keys}"
                else:
                    new_block = row['contracted_block']
                    block_note = None

                db.execute(
                    "UPDATE pickup_config SET contracted_block=?, status='active' WHERE id=?",
                    (new_block, cid)
                )
                msg = f"BID {bid} CID {cid} ({row['hotel']}): "
                parts = []
                if block_note:
                    parts.append(block_note)
                if need_unarch:
                    parts.append("unarchived")
                results.append(msg + '; '.join(parts))
            else:
                results.append(f"BID {bid} CID {cid} ({row['hotel']}): no change needed")

    db.commit()

    html = '<h3>Wrong-Year Block Fix — Results</h3>'
    if results:
        html += '<ul>' + ''.join(f'<li>{r}</li>' for r in results) + '</ul>'
    else:
        html += '<p>No matching records found.</p>'
    html += '<p><a href="/status-board">Go to Status Board</a></p>'
    return html


@app.route('/admin/fix-stray-block-dates')
def admin_fix_stray_block_dates():
    """Remove individual contracted_block date keys that fall outside a booking's
    ReportPipeline start/end window (catches same-year stray dates missed by fix-wrong-year-blocks)."""
    user = get_current_user()
    if not user or not has_permission(user, 'admin_panel'):
        return 'Forbidden — admin login required.', 403

    db      = get_db()
    results = []

    # Load all active pickup_configs that have a booking_id and a non-empty block
    rows = db.execute(
        "SELECT id, booking_id, hotel, event_name, contracted_block "
        "FROM pickup_config WHERE status != 'archived' AND booking_id IS NOT NULL "
        "AND contracted_block IS NOT NULL AND contracted_block != '{}'"
    ).fetchall()

    # Fetch ReportPipeline start/end for all relevant booking_ids
    bid_list = list({str(r['booking_id']) for r in rows})
    pipeline = {}
    if bid_list:
        ph = ','.join('?' * len(bid_list))
        for pr in db.execute(
            f"SELECT BookingId, StartDate, EndDate FROM ReportPipeline WHERE BookingId IN ({ph})",
            bid_list
        ).fetchall():
            s = _to_iso(pr['StartDate']) or ''
            e = _to_iso(pr['EndDate'])   or ''
            if len(s) == 10:
                pipeline[str(pr['BookingId'])] = {'start': s, 'end': e}

    for row in rows:
        bid   = str(row['booking_id'])
        pdates = pipeline.get(bid)
        if not pdates:
            continue

        try:
            block = json.loads(row['contracted_block'])
        except Exception:
            continue

        p_start = pdates['start']
        p_end   = pdates['end']

        # Keep only dates that fall within [start - 7 days, end + 7 days] of the pipeline window
        # (small buffer for pre/post dates that are legitimately just outside the window)
        from datetime import datetime as _dt, timedelta as _td
        try:
            window_start = (_dt.strptime(p_start, '%Y-%m-%d') - _td(days=7)).strftime('%Y-%m-%d')
            window_end   = (_dt.strptime(p_end,   '%Y-%m-%d') + _td(days=7)).strftime('%Y-%m-%d')
        except Exception:
            continue

        stray  = [k for k in block if k < window_start or k > window_end]
        if not stray:
            continue

        clean     = {k: v for k, v in block.items() if k not in stray}
        new_block = json.dumps(clean)
        db.execute(
            "UPDATE pickup_config SET contracted_block=? WHERE id=?",
            (new_block, row['id'])
        )
        results.append(
            f"BID {bid} CID {row['id']} ({row['hotel']}): "
            f"removed stray date(s) {stray} — kept {sorted(clean.keys())}"
        )

    db.commit()

    html = '<h3>Stray Block Date Fix — Results</h3>'
    if results:
        html += '<ul>' + ''.join(f'<li>{r}</li>' for r in results) + '</ul>'
        html += f'<p><strong>{len(results)} record(s) updated.</strong></p>'
    else:
        html += '<p>No stray dates found — all blocks look clean.</p>'
    html += '<p><a href="/status-board">Go to Status Board</a></p>'
    return html


# ── Status Board ──────────────────────────────────────────────────────────────

@app.route('/status-board')
def status_board():
    """Per-user action-item dashboard for pickup tracking data quality."""
    user = get_current_user()
    if not has_permission(user, 'pickups_payments'):
        flash('Access denied.', 'error')
        return redirect(url_for('index'))

    db          = get_db()
    today       = datetime.today().strftime('%Y-%m-%d')

    acct_filter = get_pickup_account_filter(user, db)

    # --- Load active ignore records for this user ---
    ignored = set()
    for row in db.execute(
        "SELECT config_id, issue_type FROM status_board_ignore "
        "WHERE user_id=? AND (ignore_until IS NULL OR ignore_until > ?)",
        (user['id'], today)
    ).fetchall():
        ignored.add((row['config_id'], row['issue_type']))

    # --- Load all non-archived pickup_config with account filter ---
    # Also exclude records whose booking is Cancelled in ReportPipeline
    base_sql = ("SELECT * FROM pickup_config WHERE status != 'archived' "
                "AND (booking_id IS NULL OR booking_id NOT IN "
                "  (SELECT BookingId FROM ReportPipeline WHERE BookingStatus='Cancelled'))")
    params   = []
    if acct_filter is None:
        pass
    elif acct_filter:
        ph = ','.join('?' * len(acct_filter))
        base_sql += f' AND organization IN ({ph})'
        params.extend(acct_filter)
    else:
        base_sql += ' AND 1=0'
    configs = db.execute(base_sql, params).fetchall()

    # --- ReportPipeline dates fallback (used when contracted_block is empty) ---
    pipeline_dates = {}  # booking_id (str) → {'start': 'YYYY-MM-DD', 'end': 'YYYY-MM-DD'}
    bid_list = [str(cfg['booking_id']) for cfg in configs if cfg['booking_id']]
    if bid_list:
        ph = ','.join('?' * len(bid_list))
        for row in db.execute(
            f"SELECT BookingId, StartDate, EndDate FROM ReportPipeline WHERE BookingId IN ({ph})",
            bid_list
        ).fetchall():
            pipeline_dates[str(row['BookingId'])] = {
                'start': _to_iso(row['StartDate']),
                'end':   _to_iso(row['EndDate']),
            }

    # --- Latest pickup_weekly per config ---
    latest_pickup = {}
    for row in db.execute(
        "SELECT config_id, MAX(report_date) as last_date FROM pickup_weekly GROUP BY config_id"
    ).fetchall():
        latest_pickup[row['config_id']] = row['last_date']

    # --- Latest contact log per config ---
    latest_contact = {}
    for row in db.execute(
        "SELECT config_id, MAX(contact_date) as last_contact FROM pickup_contact_log GROUP BY config_id"
    ).fetchall():
        latest_contact[row['config_id']] = row['last_contact']

    # --- Configs that have ANY pickup_weekly rows ---
    has_history = set()
    for row in db.execute("SELECT DISTINCT config_id FROM pickup_weekly").fetchall():
        has_history.add(row['config_id'])

    # --- Configs that have an HHR uploaded (keyed by booking_id) ---
    has_hhr = set()
    for row in db.execute(
        "SELECT DISTINCT booking_id FROM housing_history_files WHERE booking_id IS NOT NULL"
    ).fetchall():
        has_hhr.add(str(row['booking_id']))

    # --- Issue definitions: (severity, title, icon) ---
    ISSUE_META = {
        'overdue_pickup':        ('danger',  'Overdue Pickup Report',                          'bi-exclamation-circle-fill'),
        'past_cutoff_no_history':('danger',  'Past Cutoff',                                   'bi-exclamation-triangle-fill'),
        'event_ended_no_hhr':    ('danger',  'Event Ended — No HHR Uploaded',                 'bi-file-earmark-x-fill'),
        'rooming_list_due':      ('danger',  'Rooming List Due Within 15 Days',               'bi-list-check'),
        'no_recent_contact':     ('warning', 'No Hotel Contact in 21+ Days',                  'bi-telephone-x-fill'),
        'cutoff_approaching':    ('warning', 'Cutoff Approaching — Block Not Verified',        'bi-alarm-fill'),
        'uniform_block':         ('warning', 'Block Needs Verification (All Nights Identical)','bi-grid-fill'),
        'empty_block':           ('warning', 'No Contracted Block Entered',                    'bi-calendar-x-fill'),
        'missing_hotel_email':   ('info',    'Missing Hotel Contact Email',                    'bi-envelope-x-fill'),
        'missing_client_contact':('info',    'Missing Client Name or Email',                   'bi-person-x-fill'),
        'missing_cutoff':        ('info',    'No Cutoff Date Set',                             'bi-calendar-minus-fill'),
        'missing_rate':          ('info',    'No Contracted Room Rate',                        'bi-currency-dollar'),
    }

    issues_by_type = {k: [] for k in ISSUE_META}

    for cfg in configs:
        cid   = cfg['id']
        block = {}
        try:
            block = json.loads(cfg['contracted_block'] or '{}')
        except Exception:
            pass

        dates       = sorted(block.keys())
        # Priority: 1) ReportPipeline, 2) manually-set pickup_config.event_start/end, 3) contracted_block keys
        _pdates     = pipeline_dates.get(str(cfg['booking_id'] or ''), {})
        event_start = _pdates.get('start') or cfg['event_start'] or (dates[0]  if dates else None)
        event_end   = _pdates.get('end')   or cfg['event_end']   or (dates[-1] if dates else None)
        total_block = sum(block.values()) if block else 0
        block_vals  = list(block.values())

        def _issue(itype, detail, fix_url=None):
            if (cid, itype) in ignored:
                return
            issues_by_type[itype].append({
                'config_id':   cid,
                'issue_type':  itype,
                'event_name':  cfg['event_name'] or cfg['organization'],
                'organization':cfg['organization'],
                'hotel':       cfg['hotel'] or '—',
                'booking_id':  cfg['booking_id'],
                'event_start': event_start,
                'event_end':   event_end,
                'detail':      detail,
                'fix_url':     fix_url or url_for('pickup_event', cid=cid),
            })

        is_started = event_start and event_start <= today
        is_ended   = event_end   and event_end   <  today
        is_current = is_started and not is_ended

        last_rpt  = latest_pickup.get(cid)
        last_ctct = latest_contact.get(cid)
        cutoff    = (cfg['cutoff_date'] or '').strip()

        # 1. Overdue pickup report (current events only)
        if is_current:
            if not last_rpt:
                _issue('overdue_pickup',
                       'No pickup reports filed yet — event already started.',
                       url_for('pickup_event', cid=cid))
            else:
                days_since = (datetime.strptime(today, '%Y-%m-%d') -
                              datetime.strptime(last_rpt,  '%Y-%m-%d')).days
                if days_since >= 7:
                    _issue('overdue_pickup',
                           f'Last report was {days_since} days ago ({last_rpt}).',
                           url_for('pickup_event', cid=cid))

        # 2. Past cutoff with no pickup history (only while event is still ongoing)
        if cutoff and cutoff < today and cid not in has_history and not is_ended:
            _issue('past_cutoff_no_history',
                   f'Cutoff was {cutoff} — no pickup history entered yet.',
                   url_for('pickup_event', cid=cid))

        # 2b. Event ended with no HHR uploaded
        if is_ended and str(cfg['booking_id'] or '') not in has_hhr:
            _issue('event_ended_no_hhr',
                   f'Event ended {event_end} and no Housing History Report has been uploaded.',
                   url_for('import_hhr'))

        # 2c. Rooming list required and due within 15 days of cutoff
        if cfg['rooming_list_required'] and cutoff and not is_ended:
            try:
                days_to_cutoff = (datetime.strptime(cutoff, '%Y-%m-%d') -
                                  datetime.strptime(today, '%Y-%m-%d')).days
                has_rooming = db.execute(
                    'SELECT id FROM pickup_rooming_list WHERE config_id=? LIMIT 1', (cid,)
                ).fetchone()
                if days_to_cutoff <= 15 and not has_rooming:
                    _issue('rooming_list_due',
                           f'Rooming list required — cutoff is {cutoff} ({days_to_cutoff} days away). '
                           f'Client must provide names before cut-off.',
                           url_for('pickup_event', cid=cid))
            except Exception:
                pass

        # 3. No hotel contact in 21+ days (current events only)
        if is_current:
            if not last_ctct:
                _issue('no_recent_contact',
                       'No contact with hotel has ever been logged.',
                       url_for('pickup_event', cid=cid))
            else:
                days_since = (datetime.strptime(today, '%Y-%m-%d') -
                              datetime.strptime(last_ctct, '%Y-%m-%d')).days
                if days_since >= 21:
                    _issue('no_recent_contact',
                           f'Last hotel contact logged {days_since} days ago ({last_ctct}).',
                           url_for('pickup_event', cid=cid))

        # 4. Uniform block (all nights identical — needs real nightly breakdown)
        if len(block) > 1 and len(set(block_vals)) == 1 and block_vals[0] > 0:
            _issue('uniform_block',
                   f'Every night shows exactly {block_vals[0]} rooms — likely auto-filled.')

        # 5. Empty block
        if not block or total_block == 0:
            _issue('empty_block', 'contracted_block is empty or all zeros.')

        # 6. Missing hotel email
        if not (cfg['hotel_contact_email'] or '').strip():
            _issue('missing_hotel_email', 'No hotel contact email on file.')

        # 7. Missing client contact
        no_name  = not (cfg['group_contact'] or '').strip()
        no_email = not (cfg['group_contact_email'] or '').strip()
        if no_name or no_email:
            parts = []
            if no_name:  parts.append('name')
            if no_email: parts.append('email')
            _issue('missing_client_contact',
                   'Missing client ' + ' and '.join(parts) + '.')

        # 8. Missing cutoff
        if not cutoff:
            _issue('missing_cutoff', 'No cutoff date has been set.')

        # 8b. Cutoff approaching (≤30 days) but block not yet verified
        if cutoff and cutoff >= today and not is_ended:
            try:
                days_to_cutoff = (datetime.strptime(cutoff, '%Y-%m-%d') -
                                  datetime.strptime(today, '%Y-%m-%d')).days
                block_unverified = (not block or total_block == 0 or
                                    (len(block) > 1 and len(set(block_vals)) == 1))
                if days_to_cutoff <= 30 and block_unverified:
                    _issue('cutoff_approaching',
                           f'Cutoff is in {days_to_cutoff} day(s) ({cutoff}) — '
                           f'contracted block has not been verified.',
                           url_for('pickup_edit_event', cid=cid))
            except Exception:
                pass

        # 9. Missing room rate
        if not cfg['contracted_rate'] or cfg['contracted_rate'] == 0:
            _issue('missing_rate', 'No contracted room rate entered.')

    total_issues = sum(len(v) for v in issues_by_type.values())

    return render_template('status_board.html',
                           issues_by_type=issues_by_type,
                           issue_meta=ISSUE_META,
                           total_issues=total_issues,
                           today=today)


@app.route('/status-board/ignore', methods=['POST'])
def status_board_ignore():
    """Set or update an ignore record for a status board issue."""
    user = get_current_user()
    if not has_permission(user, 'pickups_payments'):
        return jsonify({'ok': False, 'error': 'Access denied'}), 403

    data       = request.get_json(silent=True) or {}
    config_id  = int(data.get('config_id', 0))
    issue_type = data.get('issue_type', '')
    mode       = data.get('mode', '1month')   # '1month' | '2months' | 'custom' | 'permanent'
    custom_dt  = data.get('custom_date', '')

    if not config_id or not issue_type:
        return jsonify({'ok': False, 'error': 'Missing params'}), 400

    today = datetime.today()
    if mode == 'permanent':
        ignore_until = None
    elif mode == '2months':
        ignore_until = (today + timedelta(days=60)).strftime('%Y-%m-%d')
    elif mode == 'custom' and custom_dt:
        ignore_until = custom_dt
    else:  # default: 1 month
        ignore_until = (today + timedelta(days=30)).strftime('%Y-%m-%d')

    db = get_db()
    db.execute(
        '''INSERT INTO status_board_ignore (user_id, config_id, issue_type, ignore_until)
           VALUES (?,?,?,?)
           ON CONFLICT(user_id, config_id, issue_type)
           DO UPDATE SET ignore_until=excluded.ignore_until,
                         created_at=datetime('now')''',
        (user['id'], config_id, issue_type, ignore_until)
    )
    db.commit()
    label = 'permanently' if ignore_until is None else f'until {ignore_until}'
    return jsonify({'ok': True, 'label': label})


@app.route('/status-board/ignore-bulk', methods=['POST'])
def status_board_ignore_bulk():
    """Bulk-ignore multiple status board items in one call."""
    user = get_current_user()
    if not has_permission(user, 'pickups_payments'):
        return jsonify({'ok': False, 'error': 'Access denied'}), 403
    data    = request.get_json(silent=True) or {}
    items   = data.get('items', [])   # list of {config_id, issue_type}
    mode    = data.get('mode', '1month')
    custom_dt = data.get('custom_date', '')
    if not items:
        return jsonify({'ok': False, 'error': 'No items provided'}), 400
    today = datetime.today()
    if mode == 'permanent':
        ignore_until = None
    elif mode == '2months':
        ignore_until = (today + timedelta(days=60)).strftime('%Y-%m-%d')
    elif mode == 'custom' and custom_dt:
        ignore_until = custom_dt
    else:
        ignore_until = (today + timedelta(days=30)).strftime('%Y-%m-%d')
    db = get_db()
    count = 0
    for item in items:
        cid = int(item.get('config_id', 0))
        itype = item.get('issue_type', '')
        if not cid or not itype:
            continue
        db.execute(
            '''INSERT INTO status_board_ignore (user_id, config_id, issue_type, ignore_until)
               VALUES (?,?,?,?)
               ON CONFLICT(user_id, config_id, issue_type)
               DO UPDATE SET ignore_until=excluded.ignore_until, created_at=datetime('now')''',
            (user['id'], cid, itype, ignore_until)
        )
        count += 1
    db.commit()
    return jsonify({'ok': True, 'ignored': count})


# ── Tasks ────────────────────────────────────────────────────────────────────

@app.route('/pickup/<int:cid>/tasks/new', methods=['POST'])
def pickup_task_new(cid):
    db   = get_db()
    user = get_current_user()
    f    = request.form
    title    = f.get('title', '').strip()
    due_date = f.get('due_date', '').strip() or None
    assigned_uid   = f.get('assigned_to_user_id', '').strip() or None
    assigned_name  = f.get('assigned_to_name', '').strip() or None
    task_notes     = f.get('task_notes', '').strip() or None
    if not title:
        flash('Task title is required.', 'error')
        return redirect(url_for('pickup_event', cid=cid))
    if assigned_uid:
        row = db.execute("SELECT name FROM Users WHERE id=?", (assigned_uid,)).fetchone()
        if row:
            assigned_name = row['name']
    db.execute(
        '''INSERT INTO pickup_tasks
           (config_id, title, due_date, assigned_to_user_id, assigned_to_name,
            created_by_user_id, created_by_name, task_notes)
           VALUES (?,?,?,?,?,?,?,?)''',
        (cid, title, due_date, assigned_uid, assigned_name,
         user['id'] if user else None, user['name'] if user else None, task_notes)
    )
    db.commit()
    flash('Task added.', 'success')
    return redirect(url_for('pickup_event', cid=cid))


@app.route('/pickup/<int:cid>/tasks/<int:tid>/complete', methods=['POST'])
def pickup_task_complete(cid, tid):
    db   = get_db()
    user = get_current_user()
    db.execute(
        "UPDATE pickup_tasks SET completed_at=datetime('now','localtime'), completed_by=? "
        "WHERE id=? AND config_id=?",
        (user['name'] if user else 'unknown', tid, cid)
    )
    db.commit()
    flash('Task marked complete.', 'success')
    return redirect(url_for('pickup_event', cid=cid))


@app.route('/pickup/<int:cid>/tasks/<int:tid>/reopen', methods=['POST'])
def pickup_task_reopen(cid, tid):
    db = get_db()
    db.execute("UPDATE pickup_tasks SET completed_at=NULL, completed_by=NULL WHERE id=? AND config_id=?", (tid, cid))
    db.commit()
    flash('Task reopened.', 'success')
    return redirect(url_for('pickup_event', cid=cid))


@app.route('/pickup/<int:cid>/tasks/<int:tid>/delete', methods=['POST'])
def pickup_task_delete(cid, tid):
    db = get_db()
    db.execute("DELETE FROM pickup_tasks WHERE id=? AND config_id=?", (tid, cid))
    db.commit()
    flash('Task deleted.', 'success')
    return redirect(url_for('pickup_event', cid=cid))


# ── Event Notes ───────────────────────────────────────────────────────────────

@app.route('/pickup/<int:cid>/notes/add', methods=['POST'])
def pickup_note_add(cid):
    db   = get_db()
    user = get_current_user()
    text = request.form.get('note_text', '').strip()
    if not text:
        flash('Note cannot be empty.', 'error')
        return redirect(url_for('pickup_event', cid=cid))
    db.execute(
        "INSERT INTO pickup_event_notes (config_id, user_id, username, note_text) VALUES (?,?,?,?)",
        (cid, user['id'] if user else None, user['name'] if user else 'unknown', text)
    )
    db.commit()
    flash('Note added.', 'success')
    return redirect(url_for('pickup_event', cid=cid))


@app.route('/pickup/<int:cid>/notes/<int:nid>/delete', methods=['POST'])
def pickup_note_delete(cid, nid):
    user = get_current_user()
    if not user or user['role'] != 'admin':
        flash('Admin access required to delete notes.', 'error')
        return redirect(url_for('pickup_event', cid=cid))
    db = get_db()
    db.execute("DELETE FROM pickup_event_notes WHERE id=? AND config_id=?", (nid, cid))
    db.commit()
    flash('Note deleted.', 'success')
    return redirect(url_for('pickup_event', cid=cid))


def _create_pickup_config_from_booking(db, bid, account, event, hotel,
                                       start_str, end_str, peak_rooms, room_rate):
    """Create a pickup_config entry for a booking if one doesn't already exist.
    Builds the contracted_block day-by-day from StartDate→EndDate at PeakRooms per night.
    Returns True if created, False if skipped (already exists or missing data).
    """
    if not bid:
        return False
    existing = db.execute(
        "SELECT id FROM pickup_config WHERE booking_id = ?", (str(bid),)
    ).fetchone()
    if existing:
        return False

    contracted_block = {}
    if start_str and end_str:
        try:
            from datetime import date as _d, timedelta as _td
            start = _d.fromisoformat(str(start_str)[:10])
            end   = _d.fromisoformat(str(end_str)[:10])
            peak  = max(1, int(float(peak_rooms))) if peak_rooms else 1
            day   = start
            while day < end:
                contracted_block[day.isoformat()] = peak
                day += _td(days=1)
        except Exception:
            pass

    try:
        rate = float(room_rate) if room_rate else None
    except Exception:
        rate = None

    db.execute('''
        INSERT INTO pickup_config
        (booking_id, organization, event_name, hotel, contracted_block, contracted_rate, status)
        VALUES (?,?,?,?,?,?,'active')
    ''', (
        str(bid),
        account or 'Unknown',
        event or None,
        hotel or None,
        json.dumps(contracted_block),
        rate,
    ))
    return True


def sync_pickup_from_pipeline(db=None):
    """Update pickup_config organization/event_name from ReportPipeline for matched booking IDs."""
    close_after = db is None
    if db is None:
        db = get_db()
    configs = db.execute(
        "SELECT id, booking_id FROM pickup_config WHERE booking_id IS NOT NULL AND booking_id != ''"
    ).fetchall()
    updated = 0
    for c in configs:
        row = db.execute(
            'SELECT AccountName, EventName FROM ReportPipeline WHERE CAST(BookingId AS INTEGER) = CAST(? AS INTEGER) LIMIT 1',
            (c['booking_id'],)
        ).fetchone()
        if row:
            db.execute(
                'UPDATE pickup_config SET organization=COALESCE(?,organization), event_name=COALESCE(?,event_name) WHERE id=?',
                (row['AccountName'] or None, row['EventName'] or None, c['id'])
            )
            if db.execute('SELECT changes()').fetchone()[0]:
                updated += 1
    db.commit()
    if close_after:
        db.close()
    return updated


def create_pickup_configs_from_pipeline(conn=None):
    """
    Create pickup_config records for future ReportPipeline bookings that don't
    have one yet. Builds a day-by-day contracted_block from StartDate → EndDate
    using PeakRooms each night. Returns the number of new records created.
    """
    from datetime import datetime as _dt, timedelta as _td
    close_after = conn is None
    if conn is None:
        conn = get_db()
    today = _dt.today().strftime('%Y-%m-%d')
    rows = conn.execute('''
        SELECT BookingId AS booking_id, BookingName AS booking_name,
               EventName AS event_name, AccountName AS org,
               StartDate AS start_date, EndDate AS end_date,
               PeakRooms AS peak_rooms, RoomRate AS room_rate, Customer AS hotel
        FROM ReportPipeline
        WHERE EndDate > ?
          AND COALESCE(BookingStatus,'') NOT IN ('Lost','Cancelled','Turned Down')
          AND PeakRooms IS NOT NULL AND PeakRooms != ''
          AND CAST(PeakRooms AS REAL) > 0
          AND NOT EXISTS (
              SELECT 1 FROM pickup_config pc
              WHERE CAST(pc.booking_id AS INTEGER) = CAST(BookingId AS INTEGER)
          )
    ''', (today,)).fetchall()
    created = 0
    for r in rows:
        try:
            start_str = (r['start_date'] or '')[:10]
            end_str   = (r['end_date']   or '')[:10]
            if not start_str or not end_str:
                continue
            start_d = _dt.strptime(start_str, '%Y-%m-%d')
            end_d   = _dt.strptime(end_str,   '%Y-%m-%d')
            peak    = int(float(r['peak_rooms']))
            rate    = float(r['room_rate']) if r['room_rate'] else None
            block   = {}
            cur = start_d
            while cur < end_d:
                block[cur.strftime('%Y-%m-%d')] = peak
                cur += _td(days=1)
            if not block:
                continue
            cutoff     = (start_d - _td(days=30)).strftime('%Y-%m-%d')
            org        = r['org'] or ''
            event_name = r['event_name'] or r['booking_name'] or org
            hotel      = r['hotel'] or ''
            booking_id = str(r['booking_id']).split('.')[0]
            conn.execute('''
                INSERT INTO pickup_config
                    (booking_id, organization, event_name, hotel, contracted_block,
                     contracted_rate, cutoff_date, attrition_pct, status,
                     hotel_contacts, cc_emails, force_current, force_past)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'active', '[]', '[]', 0, 0)
            ''', (booking_id, org, event_name, hotel,
                  json.dumps(block), rate, cutoff, 0.80))
            created += 1
        except Exception:
            continue
    if created:
        conn.commit()
    if close_after:
        conn.close()
    return created


# ── API: booking rate lookup (for pickup config form) ─────────────────────────

@app.route('/api/booking-rate')
def api_booking_rate():
    booking_id = request.args.get('booking_id', '')
    if not booking_id:
        return jsonify({'room_rate': None})
    db = get_db()
    row = db.execute(
        'SELECT RoomRate, AccountName, EventName FROM ReportPipeline WHERE CAST(BookingId AS INTEGER) = CAST(? AS INTEGER) LIMIT 1',
        (booking_id,)
    ).fetchone()
    if row:
        return jsonify({
            'room_rate':  round(float(row['RoomRate']), 2) if row['RoomRate'] else None,
            'account':    row['AccountName'] or '',
            'event_name': row['EventName'] or '',
        })
    return jsonify({'room_rate': None, 'account': '', 'event_name': ''})


# ── Pickup Tracking Routes ────────────────────────────────────────────────────

@app.route('/pickup/sync-pipeline', methods=['POST'])
def pickup_sync_pipeline():
    synced  = sync_pickup_from_pipeline()
    created = create_pickup_configs_from_pipeline()
    parts = []
    if created:
        parts.append(f'{created} new event{"s" if created != 1 else ""} created')
    if synced:
        parts.append(f'{synced} event{"s" if synced != 1 else ""} updated')
    if parts:
        flash('Sync complete: ' + ', '.join(parts) + '.', 'success')
    else:
        flash('Already up to date — nothing to sync.', 'info')
    return redirect(url_for('pickup_dashboard'))


@app.route('/pickup/fill-missing', methods=['GET', 'POST'])
def pickup_fill_missing():
    """Find future bookings that have no pickup_config entry and create them."""
    user = get_current_user()
    if not has_permission(user, 'pickups_payments'):
        flash('Access denied.', 'error')
        return redirect(url_for('index'))

    db           = get_db()
    today        = datetime.today().strftime('%Y-%m-%d')
    acct_filter  = get_pickup_account_filter(user)   # None = no filter, [] = none, [...] = allowed list

    # Base query: future non-cancelled Kristin House bookings without a pickup_config entry
    base_sql = '''
        SELECT r.BookingId, r.BookingName, r.EventName, r.AccountName,
               r.Customer, r.StartDate, r.EndDate, r.PeakRooms, r.RoomRate,
               r.BookingAssociate, r.BookingStatus
        FROM ReportPipeline r
        WHERE (r.BookingStatus IS NULL OR r.BookingStatus NOT LIKE '%Cancel%')
          AND r.EndDate >= ?
          AND r.BookingAssociate = 'Kristin House'
          AND NOT EXISTS (
              SELECT 1 FROM pickup_config p WHERE p.booking_id = CAST(r.BookingId AS TEXT)
          )
    '''
    params = [today]

    if acct_filter is None:
        # Admin — no account restriction
        pass
    elif acct_filter:
        ph = ','.join('?' * len(acct_filter))
        base_sql += f' AND r.AccountName IN ({ph})'
        params.extend(acct_filter)
    else:
        # User has no accounts assigned — show nothing
        base_sql += ' AND 1=0'

    base_sql += ' ORDER BY r.StartDate'
    missing = db.execute(base_sql, params).fetchall()

    if request.method == 'POST':
        selected_ids = request.form.getlist('booking_ids')
        created = 0
        for row in missing:
            bid = str(row['BookingId'])
            if bid not in selected_ids:
                continue
            ok = _create_pickup_config_from_booking(
                db, bid,
                account    = row['AccountName'],
                event      = row['BookingName'] or row['EventName'],
                hotel      = row['Customer'],
                start_str  = row['StartDate'],
                end_str    = row['EndDate'],
                peak_rooms = row['PeakRooms'],
                room_rate  = row['RoomRate'],
            )
            if ok:
                created += 1
        db.commit()
        flash(f'{created} pickup tracking event{"s" if created != 1 else ""} created.', 'success')
        return redirect(url_for('pickup_dashboard'))

    return render_template('pickup_fill_missing.html', missing=missing, today=today)


# ── Contact-book helpers ──────────────────────────────────────────────────────

def _upsert_contacts(db, hotel_contacts, hotel_name, client_contacts, organization):
    """Save hotel and client contacts to their respective lookup tables."""
    for c in hotel_contacts:
        name  = (c.get('name')  or '').strip()
        email = (c.get('email') or '').strip().lower()
        if not email:
            continue
        db.execute('''
            INSERT INTO hotel_contact_list (name, email, hotel_name, created_at, updated_at)
            VALUES (?, ?, ?, datetime('now'), datetime('now'))
            ON CONFLICT(email) DO UPDATE SET
                name       = excluded.name,
                hotel_name = COALESCE(excluded.hotel_name, hotel_name),
                updated_at = datetime('now')
        ''', (name, email, hotel_name or None))

    for c in client_contacts:
        name  = (c.get('name')  or '').strip()
        email = (c.get('email') or '').strip().lower()
        if not email:
            continue
        db.execute('''
            INSERT INTO client_contacts (name, email, organization, created_at, updated_at)
            VALUES (?, ?, ?, datetime('now'), datetime('now'))
            ON CONFLICT(email) DO UPDATE SET
                name         = excluded.name,
                organization = COALESCE(excluded.organization, organization),
                updated_at   = datetime('now')
        ''', (name, email, organization or None))


@app.route('/api/contacts/clients')
def api_client_contacts():
    q = request.args.get('q', '').strip()
    if len(q) < 1:
        return jsonify([])
    db = get_db()
    rows = db.execute(
        '''SELECT name, email, organization FROM client_contacts
           WHERE name LIKE ? OR email LIKE ? OR organization LIKE ?
           ORDER BY name LIMIT 20''',
        (f'%{q}%', f'%{q}%', f'%{q}%')
    ).fetchall()
    return jsonify([dict(r) for r in rows])


@app.route('/api/contacts/hotels')
def api_hotel_contacts():
    q = request.args.get('q', '').strip()
    if len(q) < 1:
        return jsonify([])
    db = get_db()
    rows = db.execute(
        '''SELECT name, email, hotel_name FROM hotel_contact_list
           WHERE name LIKE ? OR email LIKE ? OR hotel_name LIKE ?
           ORDER BY name LIMIT 20''',
        (f'%{q}%', f'%{q}%', f'%{q}%')
    ).fetchall()
    return jsonify([dict(r) for r in rows])


@app.route('/contacts')
def contacts_page():
    db = get_db()
    clients = [dict(r) for r in db.execute('SELECT * FROM client_contacts ORDER BY name').fetchall()]
    hotels  = [dict(r) for r in db.execute('SELECT * FROM hotel_contact_list ORDER BY name').fetchall()]
    return render_template('contacts.html', clients=clients, hotels=hotels)


@app.route('/contacts/delete/<string:tbl>/<int:cid>', methods=['POST'])
def contact_delete(tbl, cid):
    if tbl not in ('client', 'hotel'):
        abort(400)
    table = 'client_contacts' if tbl == 'client' else 'hotel_contact_list'
    db = get_db()
    db.execute(f'DELETE FROM {table} WHERE id=?', (cid,))
    db.commit()
    flash('Contact deleted.', 'success')
    return redirect(url_for('contacts_page') + ('?tab=hotels' if tbl == 'hotel' else ''))


_CONTACT_FIELDS = ['name', 'email', 'title', 'phone_office', 'phone_cell',
                   'fax', 'address', 'city', 'state', 'zip', 'website', 'linkedin', 'notes']


@app.route('/contacts/add/<string:tbl>', methods=['POST'])
def contact_add(tbl):
    if tbl not in ('client', 'hotel'):
        abort(400)
    db = get_db()
    f = request.form
    name  = (f.get('name')  or '').strip()
    email = (f.get('email') or '').strip().lower()
    if not name or not email:
        flash('Name and email are required.', 'danger')
        return redirect(url_for('contacts_page') + ('?tab=hotels' if tbl == 'hotel' else ''))
    if tbl == 'client':
        db.execute('''
            INSERT INTO client_contacts
              (name, email, organization, title, phone_office, phone_cell, fax,
               address, city, state, zip, website, linkedin, notes, created_at, updated_at)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,datetime('now'),datetime('now'))
            ON CONFLICT(email) DO UPDATE SET
              name=excluded.name, organization=excluded.organization,
              title=excluded.title, phone_office=excluded.phone_office,
              phone_cell=excluded.phone_cell, fax=excluded.fax,
              address=excluded.address, city=excluded.city, state=excluded.state,
              zip=excluded.zip, website=excluded.website, linkedin=excluded.linkedin,
              notes=excluded.notes, updated_at=datetime('now')
        ''', (name, email, f.get('organization','').strip() or None,
              f.get('title','').strip() or None, f.get('phone_office','').strip() or None,
              f.get('phone_cell','').strip() or None, f.get('fax','').strip() or None,
              f.get('address','').strip() or None, f.get('city','').strip() or None,
              f.get('state','').strip() or None, f.get('zip','').strip() or None,
              f.get('website','').strip() or None, f.get('linkedin','').strip() or None,
              f.get('notes','').strip() or None))
    else:
        db.execute('''
            INSERT INTO hotel_contact_list
              (name, email, hotel_name, title, phone_office, phone_cell, fax,
               address, city, state, zip, website, linkedin, notes, created_at, updated_at)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,datetime('now'),datetime('now'))
            ON CONFLICT(email) DO UPDATE SET
              name=excluded.name, hotel_name=excluded.hotel_name,
              title=excluded.title, phone_office=excluded.phone_office,
              phone_cell=excluded.phone_cell, fax=excluded.fax,
              address=excluded.address, city=excluded.city, state=excluded.state,
              zip=excluded.zip, website=excluded.website, linkedin=excluded.linkedin,
              notes=excluded.notes, updated_at=datetime('now')
        ''', (name, email, f.get('hotel_name','').strip() or None,
              f.get('title','').strip() or None, f.get('phone_office','').strip() or None,
              f.get('phone_cell','').strip() or None, f.get('fax','').strip() or None,
              f.get('address','').strip() or None, f.get('city','').strip() or None,
              f.get('state','').strip() or None, f.get('zip','').strip() or None,
              f.get('website','').strip() or None, f.get('linkedin','').strip() or None,
              f.get('notes','').strip() or None))
    db.commit()
    flash('Contact saved.', 'success')
    return redirect(url_for('contacts_page') + ('?tab=hotels' if tbl == 'hotel' else ''))


@app.route('/contacts/edit/<string:tbl>/<int:cid>', methods=['POST'])
def contact_edit(tbl, cid):
    if tbl not in ('client', 'hotel'):
        abort(400)
    db = get_db()
    f = request.form
    name  = (f.get('name')  or '').strip()
    email = (f.get('email') or '').strip().lower()
    if not name or not email:
        flash('Name and email are required.', 'danger')
        return redirect(url_for('contacts_page') + ('?tab=hotels' if tbl == 'hotel' else ''))
    extra_col = 'organization' if tbl == 'client' else 'hotel_name'
    extra_val = f.get(extra_col, '').strip() or None
    db.execute(f'''
        UPDATE {"client_contacts" if tbl=="client" else "hotel_contact_list"}
        SET name=?, email=?, {extra_col}=?, title=?, phone_office=?, phone_cell=?,
            fax=?, address=?, city=?, state=?, zip=?, website=?, linkedin=?, notes=?,
            updated_at=datetime('now')
        WHERE id=?
    ''', (name, email, extra_val,
          f.get('title','').strip() or None, f.get('phone_office','').strip() or None,
          f.get('phone_cell','').strip() or None, f.get('fax','').strip() or None,
          f.get('address','').strip() or None, f.get('city','').strip() or None,
          f.get('state','').strip() or None, f.get('zip','').strip() or None,
          f.get('website','').strip() or None, f.get('linkedin','').strip() or None,
          f.get('notes','').strip() or None, cid))
    db.commit()
    flash('Contact updated.', 'success')
    return redirect(url_for('contacts_page') + ('?tab=hotels' if tbl == 'hotel' else ''))


@app.route('/pickup/auto-match', methods=['GET', 'POST'])
def pickup_auto_match():
    """Find pickup configs missing booking_id and suggest matches from ReportPipeline."""
    user = get_current_user()
    if not has_permission(user, 'pickups_payments'):
        flash('Access denied.', 'error')
        return redirect(url_for('pickup_dashboard'))
    db = get_db()

    if request.method == 'POST':
        # Apply confirmed matches
        applied = 0
        for key, val in request.form.items():
            if key.startswith('match_') and val:
                cid = int(key.replace('match_', ''))
                db.execute('UPDATE pickup_config SET booking_id=? WHERE id=? AND (booking_id IS NULL OR booking_id="")',
                           (val, cid))
                applied += 1
        db.commit()
        flash(f'Applied {applied} booking ID match(es).', 'success')
        return redirect(url_for('pickup_dashboard'))

    # Find unmatched configs
    unmatched = db.execute(
        "SELECT * FROM pickup_config WHERE (booking_id IS NULL OR booking_id='') AND status='active'"
    ).fetchall()

    results = []
    for c in unmatched:
        org   = (c['organization'] or '').strip().lower()
        event = (c['event_name'] or '').strip().lower()
        hotel = (c['hotel'] or '').strip().lower()

        # Score each booking: +3 org match, +3 event match, +2 hotel match
        bookings = db.execute(
            "SELECT BookingId, AccountName, EventName, Customer, StartDate, EndDate "
            "FROM ReportPipeline WHERE BookingStatus NOT LIKE '%Cancel%' OR BookingStatus IS NULL"
        ).fetchall()

        scored = []
        for b in bookings:
            score = 0
            b_org   = (b['AccountName'] or '').strip().lower()
            b_event = (b['EventName'] or '').strip().lower()
            b_hotel = (b['Customer'] or '').strip().lower()
            if org   and b_org   and (org in b_org   or b_org in org):   score += 3
            if event and b_event and (event in b_event or b_event in event): score += 3
            if hotel and b_hotel and (hotel in b_hotel or b_hotel in hotel): score += 2
            if score >= 3:
                scored.append((score, dict(b)))

        scored.sort(key=lambda x: -x[0])
        results.append({'config': c, 'suggestions': scored[:5]})

    return render_template('pickup_auto_match.html', results=results)


@app.route('/pickup')
def pickup_dashboard():
    from datetime import date, timedelta
    user = get_current_user()
    db = get_db()
    today = date.today()
    today_str = today.isoformat()
    future_cutoff = today + timedelta(days=120)

    # Account-level filter (applies to all users including admins for Pickup)
    acct_filter = get_pickup_account_filter(user)
    if acct_filter is None:
        # no restriction
        acct_where = ''
        acct_params = []
    elif acct_filter:
        ph = ','.join('?' * len(acct_filter))
        acct_where = f' AND (p.AccountName IN ({ph}) OR c.organization IN ({ph}))'
        acct_params = acct_filter + acct_filter
    else:
        acct_where = ' AND 1=0'
        acct_params = []

    configs = db.execute(f"""
        SELECT c.*,
               p.StartDate  AS bk_start,
               p.EndDate    AS bk_end,
               p.EventName  AS bk_event,
               p.AccountName AS bk_org,
               p.Customer   AS bk_hotel,
               p.PeakRooms  AS bk_peak_rooms,
               p.TotalRoomNights AS bk_total_rn
        FROM pickup_config c
        LEFT JOIN ReportPipeline p ON CAST(c.booking_id AS TEXT)=CAST(p.BookingId AS TEXT)
        WHERE c.status='active'{acct_where} ORDER BY c.cutoff_date
    """, acct_params).fetchall()
    archived_configs = db.execute(f"""
        SELECT c.*,
               p.StartDate AS bk_start, p.EndDate AS bk_end
        FROM pickup_config c
        LEFT JOIN ReportPipeline p ON CAST(c.booking_id AS TEXT)=CAST(p.BookingId AS TEXT)
        WHERE c.status='archived'{acct_where} ORDER BY c.cutoff_date
    """, acct_params).fetchall()
    past_rows, current_rows, future_rows, archived_rows, missing_hhr_rows = [], [], [], [], []

    # Pre-compute which booking_ids have HHR uploaded (needed for categorization)
    has_hhr_bids = {str(r[0]) for r in db.execute(
        "SELECT DISTINCT booking_id FROM housing_history_files WHERE booking_id IS NOT NULL"
    ).fetchall()}

    for c in configs:
        last = db.execute(
            """SELECT * FROM pickup_weekly WHERE config_id=?
               AND (total_rooms IS NOT NULL AND total_rooms > 0)
               AND (label IS NULL OR (label NOT LIKE 'ASAS %' AND label NOT LIKE 'Historical%'
                                      AND label NOT LIKE '%Final%'))
               ORDER BY report_date DESC LIMIT 1""",
            (c['id'],)
        ).fetchone()
        last_any = db.execute(
            "SELECT report_date FROM pickup_weekly WHERE config_id=? ORDER BY report_date DESC LIMIT 1",
            (c['id'],)
        ).fetchone()
        all_weekly = db.execute("SELECT label FROM pickup_weekly WHERE config_id=?", (c['id'],)).fetchall()
        last_contact = db.execute(
            "SELECT contact_date, contact_type FROM pickup_contact_log WHERE config_id=? AND contact_type='email_sent' ORDER BY contact_date DESC, id DESC LIMIT 1",
            (c['id'],)
        ).fetchone()
        responded_row = db.execute(
            "SELECT contact_date FROM pickup_contact_log WHERE config_id=? AND contact_type='responded' ORDER BY contact_date DESC LIMIT 1",
            (c['id'],)
        ).fetchone()
        has_response  = bool(responded_row)
        responded_date = responded_row['contact_date'] if responded_row else None
        block = json.loads(c['contracted_block'] or '{}')
        contracted_total = sum(block.values())
        last_total = last['total_rooms'] if last else 0
        last_pct_block = last['pct_of_block'] if last else None
        last_pct_attr  = last['pct_of_attrition'] if last else None
        last_date      = last_any['report_date'] if last_any else None
        last_ota       = last['ota_rate'] if last else None
        if last_pct_block is None:
            badge, badge_label = 'secondary', 'No data'
        elif last_pct_block >= 80:
            badge, badge_label = 'success', 'On Pace'
        elif last_pct_block >= 60:
            badge, badge_label = 'warning', 'Watch'
        else:
            badge, badge_label = 'danger', 'At Risk'
        days_to_cutoff = None
        if c['cutoff_date']:
            try:
                days_to_cutoff = (date.fromisoformat(c['cutoff_date']) - today).days
            except Exception:
                pass
        has_final_history = any(w['label'] and 'final' in w['label'].lower() for w in all_weekly)
        all_dates = sorted(block.keys())
        # Priority: 1) ReportPipeline (bk_start/bk_end), 2) manually-set pickup_config.event_start/end, 3) contracted_block keys
        # bk_start may come as ISO datetime "2026-07-24T00:00:00.000Z" — strip time part
        event_start = _to_iso(c['bk_start']) or c['event_start'] or (all_dates[0]  if all_dates else None)
        event_end   = _to_iso(c['bk_end'])   or c['event_end']   or (all_dates[-1] if all_dates else None)
        force_past = bool(c['force_past']) if c['force_past'] else False
        row = {
            'config': c, 'contracted_total': contracted_total,
            'last_total': last_total, 'last_pct_block': last_pct_block,
            'last_pct_attr': last_pct_attr, 'last_date': last_date,
            'last_ota': last_ota, 'days_to_cutoff': days_to_cutoff,
            'badge': badge, 'badge_label': badge_label,
            'has_final_history': has_final_history,
            'event_start': event_start, 'event_end': event_end,
            'force_current': bool(c['force_current']),
            'force_past': force_past,
            'last_contact': last_contact,
            'has_response': has_response, 'responded_date': responded_date,
        }
        start_date = None
        for _sd in [_to_iso(c['bk_start']),
                    c['event_start'],
                    all_dates[0] if all_dates else None]:
            if _sd:
                try:
                    start_date = date.fromisoformat(_sd)
                    break
                except Exception:
                    pass
        end_date = None
        for _ed in [_to_iso(c['bk_end']),
                    c['event_end'],
                    all_dates[-1] if all_dates else None]:
            if _ed:
                try:
                    end_date = date.fromisoformat(_ed)
                    break
                except Exception:
                    pass
        meeting_ended = end_date is not None and end_date < today
        has_hhr = str(c['booking_id'] or '') in has_hhr_bids
        if force_past or (has_final_history and meeting_ended):
            past_rows.append(row)
        elif meeting_ended and has_hhr:
            past_rows.append(row)
        elif meeting_ended and not has_hhr:
            missing_hhr_rows.append(row)
        elif c['force_current']:
            current_rows.append(row)
        elif start_date is not None and start_date > future_cutoff:
            future_rows.append(row)
        else:
            current_rows.append(row)

    for c in archived_configs:
        block = json.loads(c['contracted_block'] or '{}')
        all_dates = sorted(block.keys())
        archived_rows.append({
            'config': c,
            'event_start': _to_iso(c['bk_start']) or c['event_start'] or (all_dates[0]  if all_dates else None),
            'event_end':   _to_iso(c['bk_end'])   or c['event_end']   or (all_dates[-1] if all_dates else None),
        })

    sort_mode = request.args.get('sort', 'date')
    if sort_mode == 'customer':
        _cur_sort  = lambda r: ((r['config']['organization'] or r['config']['bk_org'] or '').lower(),
                                r.get('event_start') or '')
        current_rows.sort(key=_cur_sort)
    else:
        _sort_key = lambda r: r.get('event_start') or ''
        current_rows.sort(key=_sort_key)
    _sort_key = lambda r: r.get('event_start') or ''
    past_rows.sort(key=_sort_key)
    missing_hhr_rows.sort(key=_sort_key)
    future_rows.sort(key=_sort_key)
    archived_rows.sort(key=_sort_key)

    # ── Build event groups (events with >1 hotel in the same section) ──────────
    import re as _re_group
    def _slugify(s):
        return _re_group.sub(r'[^a-z0-9]+', '-', (s or '').lower()).strip('-')

    def _build_groups(rows):
        """Return (groups_dict, group_order) for events with >1 hotel in rows."""
        from collections import defaultdict
        bucket = defaultdict(list)
        for r in rows:
            key = (r['config']['event_name'] or '').strip().lower()
            bucket[key].append(r)
        groups = {}
        group_order = []
        for key, grp_rows in bucket.items():
            if len(grp_rows) > 1:
                combined_block = sum(r['contracted_total'] for r in grp_rows)
                combined_total = sum(r['last_total'] for r in grp_rows)
                if combined_block:
                    combined_pct = round(combined_total / combined_block * 100, 1)
                else:
                    combined_pct = None
                if combined_pct is None:
                    grp_badge, grp_badge_label = 'secondary', 'No data'
                elif combined_pct >= 80:
                    grp_badge, grp_badge_label = 'success', 'On Pace'
                elif combined_pct >= 60:
                    grp_badge, grp_badge_label = 'warning', 'Watch'
                else:
                    grp_badge, grp_badge_label = 'danger', 'At Risk'
                starts = [r['event_start'] for r in grp_rows if r['event_start']]
                ends   = [r['event_end']   for r in grp_rows if r['event_end']]
                dates  = [r['last_date']   for r in grp_rows if r['last_date']]
                rep_config = grp_rows[0]['config']
                display_name = (rep_config['event_name'] or '').strip() or key
                # earliest cutoff across hotels
                cutoffs = [r['config']['cutoff_date'] for r in grp_rows if r['config']['cutoff_date']]
                groups[key] = {
                    'event_name':     display_name,
                    'organization':   rep_config['organization'] or rep_config['bk_org'] or '',
                    'hotels':         grp_rows,
                    'combined_block': combined_block,
                    'combined_total': combined_total,
                    'combined_pct':   combined_pct,
                    'event_start':    min(starts) if starts else None,
                    'event_end':      max(ends)   if ends   else None,
                    'last_date':      max(dates)  if dates  else None,
                    'cutoff_date':    min(cutoffs) if cutoffs else None,
                    'badge':          grp_badge,
                    'badge_label':    grp_badge_label,
                    'safe_id':        _slugify(display_name),
                    'primary_cid':    min(r['config']['id'] for r in grp_rows),
                }
                group_order.append(key)
        return groups, group_order

    current_groups,     current_group_order     = _build_groups(current_rows)
    future_groups,      future_group_order      = _build_groups(future_rows)
    past_groups,        past_group_order        = _build_groups(past_rows)
    missing_hhr_groups, missing_hhr_group_order = _build_groups(missing_hhr_rows)

    # ── KPI metrics ──────────────────────────────────────────────────────────
    def _sum_block(rows):
        t = 0
        for r in rows:
            try:
                t += sum(json.loads(r['config']['contracted_block'] or '{}').values())
            except Exception:
                pass
        return t

    kpi_active_events      = len(current_rows)
    kpi_future_events      = len(future_rows)
    kpi_contracted_rn      = _sum_block(current_rows) + _sum_block(future_rows)
    kpi_current_pickup     = sum((r['last_total'] or 0) for r in current_rows)
    kpi_contracted_rn_cur  = _sum_block(current_rows)
    kpi_pickup_pct         = round(kpi_current_pickup / kpi_contracted_rn_cur * 100, 1) if kpi_contracted_rn_cur else None

    # Status board open items — count from issues_by_type logic (quick query)
    kpi_status_items = db.execute(
        "SELECT COUNT(*) FROM pickup_config WHERE status='active'"
    ).fetchone()[0]  # placeholder — actual count needs status board logic; show total active

    # Events missing HHR (ended with no HHR uploaded) — now its own section
    kpi_missing_hhr = len(missing_hhr_rows)

    kpis = {
        'active_events':   kpi_active_events,
        'future_events':   kpi_future_events,
        'contracted_rn':   kpi_contracted_rn,
        'current_pickup':  kpi_current_pickup,
        'pickup_pct':      kpi_pickup_pct,
        'missing_hhr':     kpi_missing_hhr,
    }

    return render_template('pickup_dashboard.html',
                           past_rows=past_rows, current_rows=current_rows,
                           future_rows=future_rows, archived_rows=archived_rows,
                           missing_hhr_rows=missing_hhr_rows,
                           current_groups=current_groups, current_group_order=current_group_order,
                           future_groups=future_groups,   future_group_order=future_group_order,
                           past_groups=past_groups,       past_group_order=past_group_order,
                           missing_hhr_groups=missing_hhr_groups,
                           missing_hhr_group_order=missing_hhr_group_order,
                           today=today_str, sort_mode=sort_mode, kpis=kpis)


# ── Import pickup data from Excel (NCSL-style master spreadsheet) ─────────────

@app.route('/pickup/import-xlsx', methods=['GET', 'POST'])
def pickup_import_xlsx():
    if request.method == 'GET':
        return render_template('pickup_import_xlsx.html')

    f = request.files.get('xlsx_file')
    if not f or not f.filename:
        flash('Please select an Excel file.', 'warning')
        return redirect(url_for('pickup_import_xlsx'))

    file_bytes = f.read()
    try:
        from pickup_utils import parse_pickup_xlsx
        grids = parse_pickup_xlsx(file_bytes)
    except Exception as e:
        flash(f'Error parsing file: {e}', 'danger')
        return redirect(url_for('pickup_import_xlsx'))

    if not grids:
        flash('No valid hotel grids found in the file.', 'warning')
        return redirect(url_for('pickup_import_xlsx'))

    # Check which booking IDs already have pickup_config records
    db = get_db()
    existing = {}
    for g in grids:
        bid = g['booking_id']
        if bid:
            row = db.execute(
                'SELECT id, event_name, hotel FROM pickup_config WHERE booking_id=?', (bid,)
            ).fetchone()
            if row:
                existing[bid] = dict(row)

    import json as _json
    grids_json = _json.dumps(grids)
    return render_template('pickup_import_xlsx_review.html',
                           grids=grids, grids_json=grids_json, existing=existing,
                           filename=f.filename)


@app.route('/pickup/import-xlsx/confirm', methods=['POST'])
def pickup_import_xlsx_confirm():
    import json as _json
    from datetime import date
    grids = _json.loads(request.form.get('grids_json', '[]'))
    selected_indices = set(request.form.getlist('selected'))
    db = get_db()
    today = date.today().isoformat()

    created = updated = skipped = pickups_added = 0
    for i, g in enumerate(grids):
        if str(i) not in selected_indices:
            skipped += 1
            continue

        bid = g.get('booking_id') or None
        contracted_block = g.get('contracted_block') or {}
        contracted_rate  = g.get('contracted_rate')
        attrition_pct    = g.get('attrition_pct')
        pickups          = g.get('pickups', [])

        # Contacts
        hc_name   = g.get('contact_name', '') or ''
        hc_phone  = g.get('contact_phone', '') or ''
        hc_email  = g.get('contact_email', '') or ''
        gc_name   = g.get('gc_name', '') or ''
        gc_email  = g.get('gc_email', '') or ''

        # Derive cutoff_date from last date in contracted_block
        cutoff_date = None
        if contracted_block:
            cutoff_date = max(contracted_block.keys())

        # Build hotel_contacts JSON (includes phone)
        hc = []
        if hc_name or hc_email:
            entry = {'name': hc_name, 'email': hc_email}
            if hc_phone:
                entry['phone'] = hc_phone
            hc = [entry]

        # Check existing
        existing_row = None
        if bid:
            existing_row = db.execute(
                'SELECT id FROM pickup_config WHERE booking_id=?', (bid,)
            ).fetchone()

        if existing_row:
            cid = existing_row['id']
            db.execute('''UPDATE pickup_config SET
                contracted_block=?, contracted_rate=?, attrition_pct=?, cutoff_date=?,
                hotel_contact=?, hotel_contact_email=?, hotel_contacts=?,
                group_contact=?, group_contact_email=?
                WHERE id=?''',
                (_json.dumps(contracted_block), contracted_rate, attrition_pct,
                 cutoff_date,
                 hc_name or None, hc_email or None, _json.dumps(hc),
                 gc_name or None, gc_email or None,
                 cid))
            updated += 1
        else:
            db.execute('''INSERT INTO pickup_config
                (booking_id, organization, event_name, hotel, contracted_block, contracted_rate,
                 attrition_pct, cutoff_date, hotel_contact, hotel_contact_email, hotel_contacts,
                 group_contact, group_contact_email, cc_emails, status, force_current, force_past)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
                (bid, g.get('organization', 'NCSL'), g.get('event_name', ''),
                 g.get('hotel', ''),
                 _json.dumps(contracted_block), contracted_rate, attrition_pct,
                 cutoff_date,
                 hc_name or None, hc_email or None, _json.dumps(hc),
                 gc_name or None, gc_email or None,
                 '[]', 'active', 0, 0))
            cid = db.execute('SELECT last_insert_rowid()').fetchone()[0]
            created += 1

        # Insert pickup_weekly records (skip duplicates by report_date)
        existing_dates = {r['report_date'] for r in db.execute(
            'SELECT report_date FROM pickup_weekly WHERE config_id=?', (cid,)).fetchall()}

        for p in pickups:
            rd = p.get('report_date')
            if not rd or rd in existing_dates:
                continue
            db.execute('''INSERT INTO pickup_weekly
                (config_id, report_date, pickup_by_night, total_rooms, change_from_last,
                 pct_of_block, pct_of_attrition, label, notes)
                VALUES (?,?,?,?,?,?,?,?,?)''',
                (cid, rd, _json.dumps(p.get('pickup_by_night', {})),
                 p.get('total_rooms'), p.get('change_from_last'),
                 p.get('pct_of_block'), p.get('pct_of_attrition'), None, None))
            pickups_added += 1
            existing_dates.add(rd)

    db.commit()
    parts = []
    if created:   parts.append(f'{created} event{"s" if created != 1 else ""} created')
    if updated:   parts.append(f'{updated} event{"s" if updated != 1 else ""} updated')
    if skipped:   parts.append(f'{skipped} skipped')
    if pickups_added: parts.append(f'{pickups_added} pickup record{"s" if pickups_added != 1 else ""} added')
    flash('Import complete: ' + ', '.join(parts) + '.', 'success')
    return redirect(url_for('pickup_dashboard'))


@app.route('/pickup/new', methods=['GET', 'POST'])
def pickup_new_event():
    if request.method == 'POST':
        f = request.form
        contracted_block = {}
        for d, r in zip(f.getlist('block_date'), f.getlist('block_rooms')):
            if d and r:
                contracted_block[d] = int(r)
        attrition_raw = f.get('attrition_pct', '')
        attrition = float(attrition_raw) / 100 if attrition_raw else None
        ota_url = f.get('ota_url', '').strip() or None
        hc_name   = f.get('hotel_contact', '').strip() or None
        hc_email  = f.get('hotel_contact_email', '').strip() or None
        hc2_name  = f.get('hotel_contact2', '').strip() or hc_name   # default = primary
        hc2_email = f.get('hotel_contact2_email', '').strip() or hc_email
        gc_name  = f.get('group_contact', '').strip() or None
        gc_email = f.get('group_contact_email', '').strip() or None
        hotel_contacts = [{'name': hc_name or '', 'email': hc_email or ''}] if (hc_name or hc_email) else []
        # Additional CC recipients
        cc_emails = []
        for n, e in zip(f.getlist('cc_name[]'), f.getlist('cc_email[]')):
            if n.strip() or e.strip():
                cc_emails.append({'name': n.strip(), 'email': e.strip()})
        db = get_db()

        # ── Duplicate detection ──────────────────────────────────────────────
        new_bid   = (f.get('booking_id') or '').strip()
        new_event = (f.get('event_name') or '').strip().lower()
        new_hotel = (f.get('hotel') or '').strip().lower()
        duplicates = []
        if new_bid:
            bid_dups = db.execute(
                "SELECT id, event_name, hotel, status FROM pickup_config "
                "WHERE CAST(booking_id AS TEXT)=CAST(? AS TEXT)",
                (new_bid,)
            ).fetchall()
            for d in bid_dups:
                duplicates.append(
                    f"Booking ID {new_bid} already exists: "
                    f"<strong>{d['event_name']}</strong> @ {d['hotel']} "
                    f"({'archived' if d['status']=='archived' else 'active'}) — "
                    f"<a href=\"{url_for('pickup_event', cid=d['id'])}\">View</a>"
                )
        if new_event and new_hotel:
            name_dups = db.execute(
                "SELECT id, event_name, hotel, status FROM pickup_config "
                "WHERE LOWER(TRIM(event_name))=? AND LOWER(TRIM(hotel))=? "
                "AND (? = '' OR CAST(booking_id AS TEXT) != ?)",
                (new_event, new_hotel, new_bid, new_bid)
            ).fetchall()
            for d in name_dups:
                duplicates.append(
                    f"Same event name + hotel already exists: "
                    f"<strong>{d['event_name']}</strong> @ {d['hotel']} "
                    f"({'archived' if d['status']=='archived' else 'active'}) — "
                    f"<a href=\"{url_for('pickup_event', cid=d['id'])}\">View</a>"
                )
        if duplicates and not f.get('confirm_duplicate'):
            # Re-render the form with warnings — user must tick confirm to proceed
            pipeline_rate = pipeline_org = pipeline_event = None
            if new_bid:
                row = db.execute(
                    'SELECT RoomRate, AccountName, EventName FROM ReportPipeline WHERE CAST(BookingId AS INTEGER)=CAST(? AS INTEGER) LIMIT 1',
                    (new_bid,)
                ).fetchone()
                if row:
                    pipeline_rate  = round(float(row['RoomRate']), 2) if row['RoomRate'] else None
                    pipeline_org   = row['AccountName'] or None
                    pipeline_event = row['EventName'] or None
            return render_template('pickup_config_form.html', config=None,
                                   action=url_for('pickup_new_event'),
                                   cancel_url=url_for('pickup_dashboard'),
                                   pipeline_rate=pipeline_rate,
                                   pipeline_org=pipeline_org,
                                   pipeline_event=pipeline_event,
                                   duplicate_warnings=duplicates,
                                   form_data=f)

        db.execute('''
            INSERT INTO pickup_config
            (booking_id, tab_name, organization, event_name, hotel, hotel_contact,
             hotel_contact_email, hotel_contact2, hotel_contact2_email, hotel_contacts,
             group_contact, group_contact_email,
             cutoff_date, attrition_pct, contracted_block, contracted_rate, shoulder_pre,
             shoulder_post, hotel_booking_link, notes, ota_url, cc_emails)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        ''', (
            f.get('booking_id'), f.get('tab_name'), f['organization'],
            f.get('event_name'), f.get('hotel'), hc_name,
            hc_email, hc2_name, hc2_email, json.dumps(hotel_contacts), gc_name,
            gc_email, f.get('cutoff_date'),
            attrition, json.dumps(contracted_block),
            float(f['contracted_rate']) if f.get('contracted_rate') else None,
            int(f.get('shoulder_pre', 3)), int(f.get('shoulder_post', 3)),
            f.get('hotel_booking_link'), f.get('notes'), ota_url, json.dumps(cc_emails)
        ))
        db.commit()
        _upsert_contacts(db, hotel_contacts, f.get('hotel', ''), cc_emails, f['organization'])
        db.commit()
        flash('Event created.', 'success')
        return redirect(url_for('pickup_dashboard'))
    pipeline_rate = pipeline_org = pipeline_event = None
    booking_id_qs = request.args.get('booking_id')
    if booking_id_qs:
        db = get_db()
        row = db.execute(
            'SELECT RoomRate, AccountName, EventName FROM ReportPipeline WHERE CAST(BookingId AS INTEGER) = CAST(? AS INTEGER) LIMIT 1',
            (booking_id_qs,)
        ).fetchone()
        if row:
            if row['RoomRate']:
                pipeline_rate = round(float(row['RoomRate']), 2)
            pipeline_org   = row['AccountName'] or None
            pipeline_event = row['EventName'] or None
    return render_template('pickup_config_form.html', config=None,
                           action=url_for('pickup_new_event'),
                           cancel_url=url_for('pickup_dashboard'),
                           pipeline_rate=pipeline_rate,
                           pipeline_org=pipeline_org,
                           pipeline_event=pipeline_event)


@app.route('/pickup/<int:cid>')
def pickup_event(cid):
    from datetime import datetime as _dt, date as _date
    db = get_db()
    config = db.execute("SELECT * FROM pickup_config WHERE id=?", (cid,)).fetchone()
    if not config:
        flash('Event not found.', 'error')
        return redirect(url_for('pickup_dashboard'))
    all_weekly = db.execute("SELECT * FROM pickup_weekly WHERE config_id=? ORDER BY report_date DESC", (cid,)).fetchall()
    # Most recent real entry for OTA rate comparison (excluding historical/placeholder rows)
    last_real = db.execute(
        """SELECT * FROM pickup_weekly WHERE config_id=?
           AND (total_rooms IS NOT NULL AND total_rooms > 0)
           AND (label IS NULL OR (label NOT LIKE 'ASAS %' AND label NOT LIKE 'Historical%'
                                  AND label NOT LIKE '%Final%'))
           ORDER BY report_date DESC LIMIT 1""",
        (cid,)
    ).fetchone()
    rooming = db.execute(
        "SELECT id, upload_date, filename, total_guests, reconciliation_status, discrepancy_notes FROM pickup_rooming_list WHERE config_id=? ORDER BY upload_date DESC", (cid,)
    ).fetchall()
    contact_log = db.execute(
        "SELECT * FROM pickup_contact_log WHERE config_id=? ORDER BY contact_date DESC, id DESC", (cid,)
    ).fetchall()
    block = json.loads(config['contracted_block'] or '{}')
    # Include shoulder nights from any weekly entry that fall outside the contracted block
    all_date_set = set(block.keys())
    for _w in all_weekly:
        try:
            _pbn = json.loads(_w['pickup_by_night'] or '{}')
            all_date_set.update(d for d in _pbn if d != 'historical_total')
        except Exception:
            pass
    all_dates = sorted(all_date_set)
    contracted_total = sum(v for v in block.values() if v)
    day_map = {}
    for d in all_dates:
        try:
            day_map[d] = _dt.strptime(d, '%Y-%m-%d').strftime('%a')
        except Exception:
            day_map[d] = ''
    attrition_pct   = config['attrition_pct'] or 0
    attrition_rooms = round(contracted_total * attrition_pct, 1)

    # Split rows into historical pace rows (label = 'ASAS YYYY') vs current reporting rows
    raw_historical = [w for w in all_weekly if w['label'] and (
        w['label'].startswith('Historical') or w['label'].startswith('ASAS ')
    )]
    raw_current    = [w for w in all_weekly if not (w['label'] and (
        w['label'].startswith('Historical') or w['label'].startswith('ASAS ')
    ))]

    # Build historical display list (for backward-compat; sorted most recent year first)
    historical_years = []
    for w in sorted(raw_historical, key=lambda x: x['report_date'], reverse=True):
        pbn = json.loads(w['pickup_by_night'] or '{}')
        total = pbn.get('historical_total', w['total_rooms'] or 0)
        row = dict(w)
        row['total_rooms'] = total
        historical_years.append(row)

    # Build pace comparison structure for multi-year chart
    pace_comparison = None
    if raw_historical:
        # Event start dates for ASAS historical years (confirmed from Excel data)
        ASAS_EVENT_STARTS = {
            2021: '2021-07-11',  # Louisville, KY
            2022: '2022-06-23',  # OKC
            2023: '2023-07-15',  # Albuquerque
            2024: '2024-07-17',  # Calgary
            2025: '2025-06-30',  # The Diplomat
            2026: '2026-07-17',  # Madison
        }
        from datetime import date as _ddate

        def _weeks_out(event_start_str, report_date_str):
            try:
                es = _ddate.fromisoformat(event_start_str)
                rd = _ddate.fromisoformat(report_date_str)
                return (es - rd).days // 7
            except Exception:
                return None

        # Build per-year dict: year -> list of {weeks_out, total_rooms}
        hist_by_year = {}
        for w in raw_historical:
            lbl = w['label'] or ''
            # Support both 'ASAS YYYY' and legacy 'Historical YYYY...' labels
            import re as _re
            m = _re.search(r'(\d{4})', lbl)
            if not m:
                continue
            yr = int(m.group(1))
            es = ASAS_EVENT_STARTS.get(yr)
            if not es:
                continue
            wo = _weeks_out(es, w['report_date'])
            if wo is None or wo < 0:
                continue
            pbn = json.loads(w['pickup_by_night'] or '{}')
            total = pbn.get('historical_total', w['total_rooms'] or 0)
            if yr not in hist_by_year:
                hist_by_year[yr] = {}
            # Keep the entry with highest weeks_out match; for duplicate weeks, keep most rooms
            if wo not in hist_by_year[yr] or total > hist_by_year[yr][wo]:
                hist_by_year[yr][wo] = total

        # Add 2026 current data
        current_event_start = ASAS_EVENT_STARTS.get(2026)
        if current_event_start and raw_current:
            hist_by_year[2026] = {}
            for w in raw_current:
                pbn = json.loads(w['pickup_by_night'] or '{}')
                total = sum(v for v in pbn.values() if isinstance(v, (int, float)))
                wo = _weeks_out(current_event_start, w['report_date'])
                if wo is None or wo < 0:
                    continue
                if wo not in hist_by_year[2026] or total > hist_by_year[2026][wo]:
                    hist_by_year[2026][wo] = total

        # Collect all unique weeks_out values and sort descending (most weeks out first)
        all_weeks = sorted(
            set(wo for yr_data in hist_by_year.values() for wo in yr_data.keys()),
            reverse=True
        )

        # Build by_weeks_out lookup
        by_weeks_out = {}
        for wo in all_weeks:
            by_weeks_out[wo] = {}
            for yr, yr_data in hist_by_year.items():
                if wo in yr_data:
                    by_weeks_out[wo][str(yr)] = yr_data[wo]

        years_available = sorted(hist_by_year.keys(), reverse=True)

        pace_comparison = {
            'years': years_available,
            'event_starts': ASAS_EVENT_STARTS,
            'by_weeks_out': by_weeks_out,
            'all_weeks': all_weeks,
        }

    # Build current weekly display with pct_of_block / WoW calculations
    weekly_display = []
    prev_total = None
    for w in reversed(raw_current):
        pbn = json.loads(w['pickup_by_night'] or '{}')
        computed_total = sum(v for v in pbn.values() if v is not None and v != '')
        pct_blk  = round(computed_total / contracted_total * 100, 1) if contracted_total else None
        pct_attr = round(computed_total / attrition_rooms * 100, 1) if attrition_rooms else None
        wow      = (computed_total - prev_total) if prev_total is not None else None
        row = dict(w)
        row['total_rooms']      = computed_total
        row['pct_of_block']     = pct_blk
        row['pct_of_attrition'] = pct_attr
        row['change_from_last'] = wow
        weekly_display.append(row)
        prev_total = computed_total
    weekly_display.reverse()

    past_cutoff = False
    meeting_ended = False
    has_final_history = any(w['label'] and 'final' in w['label'].lower() for w in all_weekly)
    if all_dates:
        try:
            if _dt.strptime(all_dates[0], '%Y-%m-%d') < _dt.today():
                past_cutoff = True
        except Exception:
            pass
        try:
            if _dt.strptime(all_dates[-1], '%Y-%m-%d') < _dt.today():
                meeting_ended = True
        except Exception:
            pass
    hhr_row = db.execute(
        'SELECT id FROM housing_history_files WHERE booking_id=? ORDER BY id DESC LIMIT 1',
        (config['booking_id'],)
    ).fetchone() if config['booking_id'] else None
    has_hhr = bool(hhr_row)

    # Check if this is part of a multi-hotel event
    sibling_count = db.execute(
        "SELECT COUNT(*) FROM pickup_config WHERE LOWER(TRIM(event_name))=LOWER(TRIM(?)) AND status='active'",
        ((config['event_name'] or ''),)
    ).fetchone()[0]
    is_multi_hotel = sibling_count > 1
    # Find primary_cid for the event report link
    if is_multi_hotel:
        primary_cid_for_report = db.execute(
            "SELECT MIN(id) FROM pickup_config WHERE LOWER(TRIM(event_name))=LOWER(TRIM(?)) AND status='active'",
            ((config['event_name'] or ''),)
        ).fetchone()[0]
    else:
        primary_cid_for_report = None

    # Tasks for this event
    tasks = db.execute(
        "SELECT * FROM pickup_tasks WHERE config_id=? ORDER BY completed_at IS NOT NULL, due_date ASC NULLS LAST, id DESC",
        (cid,)
    ).fetchall()

    # Notes for this event
    event_notes = db.execute(
        "SELECT * FROM pickup_event_notes WHERE config_id=? ORDER BY id DESC",
        (cid,)
    ).fetchall()

    # All active users (for task assignment)
    all_users = db.execute(
        "SELECT id, name FROM Users WHERE active=1 ORDER BY name"
    ).fetchall()

    # Recently deleted weekly reports (kept 90 days, restorable)
    deleted_weekly = db.execute(
        """SELECT * FROM pickup_weekly_deleted
           WHERE config_id=?
             AND deleted_at >= datetime('now','-90 days','localtime')
           ORDER BY deleted_at DESC""",
        (cid,)
    ).fetchall()

    last_ota_rate = last_real['ota_rate'] if last_real else None
    last_ota_url  = last_real['ota_url']  if last_real else None
    show_rate_issue = bool(
        last_ota_rate and config['contracted_rate'] and
        float(last_ota_rate) < float(config['contracted_rate'])
    )
    return render_template('pickup_event.html',
                           config=config, weekly=weekly_display, historical_years=historical_years,
                           pace_comparison=pace_comparison,
                           rooming=rooming,
                           contact_log=contact_log, block=block, all_dates=all_dates,
                           day_map=day_map, contracted_total=contracted_total,
                           attrition_pct=attrition_pct, attrition_rooms=attrition_rooms,
                           past_cutoff=past_cutoff, has_final_history=has_final_history,
                           meeting_ended=meeting_ended,
                           has_hhr=has_hhr, today=_date.today().isoformat(),
                           is_multi_hotel=is_multi_hotel,
                           primary_cid_for_report=primary_cid_for_report,
                           deleted_weekly=deleted_weekly,
                           tasks=tasks, event_notes=event_notes, all_users=all_users,
                           last_ota_rate=last_ota_rate, last_ota_url=last_ota_url,
                           show_rate_issue=show_rate_issue)


# ── Event Report: combined view across all hotels for an event ────────────────

@app.route('/pickup/event-report/<int:primary_cid>')
def pickup_event_report(primary_cid):
    """Show combined pickup data for all hotels sharing the same event_name."""
    from datetime import date as _date2, datetime as _dt2
    db = get_db()

    # Get the primary config to find the event_name
    primary = db.execute("SELECT * FROM pickup_config WHERE id=?", (primary_cid,)).fetchone()
    if not primary:
        flash('Event not found.', 'error')
        return redirect(url_for('pickup_dashboard'))

    event_name = (primary['event_name'] or '').strip()
    if not event_name:
        return redirect(url_for('pickup_event', cid=primary_cid))

    # Find all configs with the same event_name (case-insensitive, active only)
    all_configs = db.execute(
        "SELECT * FROM pickup_config WHERE LOWER(TRIM(event_name))=LOWER(TRIM(?)) AND status='active' ORDER BY id",
        (event_name,)
    ).fetchall()

    if len(all_configs) <= 1:
        # Only one hotel — redirect to normal event page
        return redirect(url_for('pickup_event', cid=primary_cid))

    organization = primary['organization'] or primary['event_name'] or ''

    # ── Per-hotel summaries ────────────────────────────────────────────────────
    hotel_summaries = []
    combined_block = 0
    all_date_union = set()

    for c in all_configs:
        block = json.loads(c['contracted_block'] or '{}')
        contracted_total = sum(block.values())
        combined_block += contracted_total
        all_date_union.update(block.keys())
        # Also include any shoulder nights from weekly entries
        for _w in db.execute("SELECT pickup_by_night FROM pickup_weekly WHERE config_id=?", (c['id'],)).fetchall():
            try:
                _pbn = json.loads(_w['pickup_by_night'] or '{}')
                all_date_union.update(d for d in _pbn if d != 'historical_total')
            except Exception:
                pass

        # Latest non-historical weekly entry (label IS NULL counts as current)
        last = db.execute(
            """SELECT * FROM pickup_weekly WHERE config_id=?
               AND (label IS NULL OR (label NOT LIKE 'ASAS %' AND label NOT LIKE 'Historical%'))
               ORDER BY report_date DESC LIMIT 1""",
            (c['id'],)
        ).fetchone()
        last_total = last['total_rooms'] if last else 0
        last_pct = last['pct_of_block'] if last else None
        last_date = last['report_date'] if last else None

        if last_pct is None:
            h_badge, h_badge_label = 'secondary', 'No data'
        elif last_pct >= 80:
            h_badge, h_badge_label = 'success', 'On Pace'
        elif last_pct >= 60:
            h_badge, h_badge_label = 'warning', 'Watch'
        else:
            h_badge, h_badge_label = 'danger', 'At Risk'

        hotel_summaries.append({
            'cid':              c['id'],
            'hotel_name':       c['hotel'] or c['event_name'] or '—',
            'contracted_total': contracted_total,
            'last_total':       last_total,
            'last_pct':         last_pct,
            'last_date':        last_date,
            'badge':            h_badge,
            'badge_label':      h_badge_label,
            'cutoff_date':      c['cutoff_date'],
        })

    all_dates = sorted(all_date_union)

    # ── Combined weekly timeline ───────────────────────────────────────────────
    # Gather all non-historical weekly rows across all hotels
    from collections import defaultdict
    date_totals = defaultdict(int)   # report_date -> sum of total_rooms
    date_hotels = defaultdict(set)   # report_date -> set of cids that reported

    for c in all_configs:
        rows_w = db.execute(
            """SELECT * FROM pickup_weekly WHERE config_id=?
               AND (label IS NULL OR (label NOT LIKE 'ASAS %' AND label NOT LIKE 'Historical%'))
               ORDER BY report_date""",
            (c['id'],)
        ).fetchall()
        for w in rows_w:
            total = w['total_rooms'] or 0
            date_totals[w['report_date']] += total
            date_hotels[w['report_date']].add(c['id'])

    # Build sorted weekly display with WoW change
    sorted_dates = sorted(date_totals.keys())
    combined_weekly = []
    prev_total = None
    for rd in sorted_dates:
        total = date_totals[rd]
        pct = round(total / combined_block * 100, 1) if combined_block else None
        wow = (total - prev_total) if prev_total is not None else None
        combined_weekly.append({
            'report_date':   rd,
            'total_rooms':   total,
            'pct_of_block':  pct,
            'wow_change':    wow,
            'hotels_count':  len(date_hotels[rd]),
        })
        prev_total = total
    combined_weekly.reverse()  # newest first

    # ── Historical pace comparison (if any hotel has ASAS historical rows) ────
    pace_comparison = None
    ASAS_EVENT_STARTS = {
        2021: '2021-07-11',
        2022: '2022-06-23',
        2023: '2023-07-15',
        2024: '2024-07-17',
        2025: '2025-06-30',
        2026: '2026-07-17',
    }

    hist_rows_all = []
    for c in all_configs:
        hist_rows = db.execute(
            """SELECT * FROM pickup_weekly WHERE config_id=?
               AND (label LIKE 'ASAS %' OR label LIKE 'Historical%')
               ORDER BY report_date""",
            (c['id'],)
        ).fetchall()
        hist_rows_all.extend(hist_rows)

    if hist_rows_all:
        import re as _re2
        def _weeks_out2(event_start_str, report_date_str):
            try:
                es = _date2.fromisoformat(event_start_str)
                rd = _date2.fromisoformat(report_date_str)
                return (es - rd).days // 7
            except Exception:
                return None

        hist_by_year = {}
        for w in hist_rows_all:
            lbl = w['label'] or ''
            m = _re2.search(r'(\d{4})', lbl)
            if not m:
                continue
            yr = int(m.group(1))
            es = ASAS_EVENT_STARTS.get(yr)
            if not es:
                continue
            wo = _weeks_out2(es, w['report_date'])
            if wo is None or wo < 0:
                continue
            pbn = json.loads(w['pickup_by_night'] or '{}')
            total = pbn.get('historical_total', w['total_rooms'] or 0)
            if yr not in hist_by_year:
                hist_by_year[yr] = {}
            if wo not in hist_by_year[yr] or total > hist_by_year[yr][wo]:
                hist_by_year[yr][wo] = total

        # Add current year from combined_weekly (reversed to get chronological)
        current_event_start = ASAS_EVENT_STARTS.get(2026)
        if current_event_start and date_totals:
            hist_by_year[2026] = {}
            for rd, total in date_totals.items():
                wo = _weeks_out2(current_event_start, rd)
                if wo is None or wo < 0:
                    continue
                if wo not in hist_by_year[2026] or total > hist_by_year[2026][wo]:
                    hist_by_year[2026][wo] = total

        if hist_by_year:
            all_weeks_hist = sorted(
                set(wo for yr_data in hist_by_year.values() for wo in yr_data.keys()),
                reverse=True
            )
            by_weeks_out = {}
            for wo in all_weeks_hist:
                by_weeks_out[wo] = {}
                for yr, yr_data in hist_by_year.items():
                    if wo in yr_data:
                        by_weeks_out[wo][str(yr)] = yr_data[wo]
            years_available = sorted(hist_by_year.keys(), reverse=True)
            pace_comparison = {
                'years':        years_available,
                'event_starts': ASAS_EVENT_STARTS,
                'by_weeks_out': by_weeks_out,
                'all_weeks':    all_weeks_hist,
            }

    # ── Cutoff range ──────────────────────────────────────────────────────────
    cutoffs = [c['cutoff_date'] for c in all_configs if c['cutoff_date']]
    cutoff_earliest = min(cutoffs) if cutoffs else None
    cutoff_latest   = max(cutoffs) if cutoffs else None

    return render_template('pickup_event_report.html',
                           event_name=event_name,
                           organization=organization,
                           configs=all_configs,
                           combined_block=combined_block,
                           combined_weekly=combined_weekly,
                           hotel_summaries=hotel_summaries,
                           pace_comparison=pace_comparison,
                           all_dates=all_dates,
                           cutoff_earliest=cutoff_earliest,
                           cutoff_latest=cutoff_latest,
                           primary_cid=primary_cid)


@app.route('/pickup/event-report/<int:primary_cid>/download-xlsx')
def pickup_event_report_xlsx(primary_cid):
    """Download a combined XLSX pace report for all hotels sharing the same event_name."""
    import io
    import re as _re_fn
    from datetime import date as _date_xlsx
    from collections import defaultdict
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    db = get_db()
    primary = db.execute("SELECT * FROM pickup_config WHERE id=?", (primary_cid,)).fetchone()
    if not primary:
        flash('Event not found.', 'error')
        return redirect(url_for('pickup_dashboard'))

    event_name = (primary['event_name'] or '').strip()
    if not event_name:
        return redirect(url_for('pickup_event', cid=primary_cid))

    organization = primary['organization'] or primary['event_name'] or ''

    # Find all active configs with the same event_name
    all_configs = db.execute(
        "SELECT * FROM pickup_config WHERE LOWER(TRIM(event_name))=LOWER(TRIM(?)) AND status='active' ORDER BY id",
        (event_name,)
    ).fetchall()

    ASAS_EVENT_STARTS = {
        2021: '2021-07-11',
        2022: '2022-06-23',
        2023: '2023-07-15',
        2024: '2024-07-17',
        2025: '2025-06-30',
        2026: '2026-07-17',
    }

    def weeks_out(event_start_str, report_date_str):
        try:
            es = _date_xlsx.fromisoformat(event_start_str)
            rd = _date_xlsx.fromisoformat(report_date_str)
            return (es - rd).days // 7
        except Exception:
            return None

    # ── Per-hotel data ─────────────────────────────────────────────────────────
    hotel_data = []
    combined_block = 0
    all_event_date_set = set()

    for c in all_configs:
        block = json.loads(c['contracted_block'] or '{}')
        contracted_total = sum(block.values())
        combined_block += contracted_total
        all_event_date_set.update(block.keys())

        weekly_rows = db.execute(
            """SELECT * FROM pickup_weekly WHERE config_id=?
               AND (label IS NULL OR (label NOT LIKE 'ASAS %' AND label NOT LIKE 'Historical%'))
               ORDER BY report_date""",
            (c['id'],)
        ).fetchall()

        # Also include any shoulder nights recorded in pickup_by_night
        for w in weekly_rows:
            try:
                pbn = json.loads(w['pickup_by_night'] or '{}')
                for d in pbn.keys():
                    if d:
                        all_event_date_set.add(d)
            except Exception:
                pass

        hotel_data.append({
            'config': c,
            'hotel_name': c['hotel'] or c['event_name'] or '—',
            'contracted_block': block,
            'contracted_total': contracted_total,
            'contracted_rate': c['contracted_rate'] or 0,
            'attrition_pct': c['attrition_pct'],
            'weekly': [dict(w) for w in weekly_rows],
        })

    event_dates = sorted(all_event_date_set)
    n_nights = len(event_dates)
    n_hotels = len(all_configs)

    # ── Historical rows across all configs ────────────────────────────────────
    import re as _re2
    hist_rows_all = []
    for c in all_configs:
        rows = db.execute(
            """SELECT * FROM pickup_weekly WHERE config_id=?
               AND (label LIKE 'ASAS %' OR label LIKE 'Historical%')
               ORDER BY report_date""",
            (c['id'],)
        ).fetchall()
        hist_rows_all.extend(rows)

    hist_by_year = {}
    for w in hist_rows_all:
        lbl = w['label'] or ''
        m = _re2.search(r'(\d{4})', lbl)
        if not m:
            continue
        yr = int(m.group(1))
        es = ASAS_EVENT_STARTS.get(yr)
        if not es:
            continue
        wo = weeks_out(es, w['report_date'])
        if wo is None or wo < 0:
            continue
        pbn = json.loads(w['pickup_by_night'] or '{}')
        total = pbn.get('historical_total', w['total_rooms'] or 0)
        if yr not in hist_by_year:
            hist_by_year[yr] = {}
        if wo not in hist_by_year[yr] or total > hist_by_year[yr][wo]:
            hist_by_year[yr][wo] = total

    # ── Build per-date aggregate (Tab 2 data) ─────────────────────────────────
    all_report_date_set = set()
    for hd in hotel_data:
        for w in hd['weekly']:
            all_report_date_set.add(w['report_date'])
    all_report_dates = sorted(all_report_date_set)

    # aggregate_by_date: report_date -> sum of total_rooms (for 2026 hist)
    aggregate_by_date = defaultdict(int)
    total_by_night = {}       # report_date -> {event_date: int}
    total_rooms_agg = {}      # report_date -> int
    hotel_pbn = {}            # cid -> report_date -> {event_date: int}

    for hd in hotel_data:
        cid = hd['config']['id']
        hotel_pbn[cid] = {}
        for w in hd['weekly']:
            rd = w['report_date']
            pbn = json.loads(w['pickup_by_night'] or '{}')
            hotel_pbn[cid][rd] = pbn
            aggregate_by_date[rd] += (w['total_rooms'] or 0)
            if rd not in total_by_night:
                total_by_night[rd] = defaultdict(int)
            for ed, rooms in pbn.items():
                if isinstance(rooms, (int, float)):
                    total_by_night[rd][ed] += int(rooms)

    for rd in all_report_dates:
        total_rooms_agg[rd] = aggregate_by_date[rd]

    # Add 2026 current to hist_by_year
    current_event_start_2026 = ASAS_EVENT_STARTS.get(2026)
    if current_event_start_2026 and aggregate_by_date:
        hist_by_year[2026] = {}
        for rd, total in aggregate_by_date.items():
            wo = weeks_out(current_event_start_2026, rd)
            if wo is None or wo < 0:
                continue
            if wo not in hist_by_year[2026] or total > hist_by_year[2026][wo]:
                hist_by_year[2026][wo] = total

    all_weeks_hist = sorted(
        set(wo for yr_data in hist_by_year.values() for wo in yr_data.keys()),
        reverse=True
    )
    years_available = sorted(hist_by_year.keys(), reverse=True)

    # ── Cutoff info ───────────────────────────────────────────────────────────
    cutoffs = [c['cutoff_date'] for c in all_configs if c['cutoff_date']]
    cutoff_earliest = min(cutoffs) if cutoffs else None

    # ── Build workbook ────────────────────────────────────────────────────────
    wb = Workbook()

    # ════════════════════════════════════════════════════════════════════════
    # TAB 1 — Historical Comparison
    # ════════════════════════════════════════════════════════════════════════
    ws1 = wb.active
    ws1.title = "Historical Comparison"

    header_fill = PatternFill("solid", fgColor="1A3A5C")
    header_font = Font(color="FFFFFF", bold=True)
    center_align = Alignment(horizontal="center")
    alt_fill = PatternFill("solid", fgColor="EAF0FB")

    # A1: event name
    ws1['A1'] = event_name
    ws1['A1'].font = Font(bold=True, size=14)
    ws1.merge_cells('A1:G1')

    # A2: organization
    ws1['A2'] = organization
    ws1['A2'].font = Font(italic=True, size=10)

    # A3: subtitle
    ws1['A3'] = "All Hotels Combined — Historical Pickup by Weeks Out"
    ws1['A3'].font = Font(size=10)

    # A4: blank
    ws1['A4'] = ""

    # Row 5: headers
    ws1.cell(5, 1).value = "Weeks Out"
    ws1.cell(5, 1).font = header_font
    ws1.cell(5, 1).fill = header_fill
    ws1.cell(5, 1).alignment = center_align

    for ci, yr in enumerate(years_available, start=2):
        lbl = f"{yr} (All Hotels)" if yr == 2026 else str(yr)
        cell = ws1.cell(5, ci)
        cell.value = lbl
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_align

    # Data rows
    for ri, wo in enumerate(all_weeks_hist):
        row_num = 6 + ri
        is_even = (ri % 2 == 0)
        row_fill = alt_fill if is_even else None

        wk_cell = ws1.cell(row_num, 1)
        wk_cell.value = f"{wo} wks"
        wk_cell.alignment = center_align
        if row_fill:
            wk_cell.fill = row_fill

        for ci, yr in enumerate(years_available, start=2):
            val = hist_by_year.get(yr, {}).get(wo)
            cell = ws1.cell(row_num, ci)
            cell.value = val if val is not None else "—"
            if val is not None:
                cell.number_format = "#,##0"
            cell.alignment = center_align
            if yr == 2026:
                cell.font = Font(bold=True)
            if row_fill:
                cell.fill = row_fill

    # Column widths Tab 1
    ws1.column_dimensions['A'].width = 12
    for ci in range(2, 2 + len(years_available)):
        ws1.column_dimensions[get_column_letter(ci)].width = 16

    # ════════════════════════════════════════════════════════════════════════
    # TAB 2 — Pace 2026
    # ════════════════════════════════════════════════════════════════════════
    ws2 = wb.create_sheet("Pace 2026")

    start_col = 5   # col E
    end_col = start_col + n_nights - 1
    total_col = end_col + 1
    pct_col = total_col + 1
    last_col_idx = pct_col + 1

    end_col_letter = get_column_letter(end_col)
    total_col_letter = get_column_letter(total_col)
    pct_col_letter = get_column_letter(pct_col)

    # ── Header block (rows 2-5) ───────────────────────────────────────────
    # Row 2: event name
    ws2.merge_cells(start_row=2, start_column=4, end_row=2, end_column=total_col)
    ws2.cell(2, 4).value = event_name
    ws2.cell(2, 4).font = Font(bold=True, size=18)
    ws2.cell(2, 4).alignment = Alignment(horizontal="center")

    # Row 3: date range
    if event_dates:
        try:
            d0 = _date_xlsx.fromisoformat(event_dates[0])
            d1 = _date_xlsx.fromisoformat(event_dates[-1])
            date_range_str = f"{d0.strftime('%B')} {d0.day}–{d1.day}, {d1.year}"
        except Exception:
            date_range_str = f"{event_dates[0]} to {event_dates[-1]}"
    else:
        date_range_str = ""
    ws2.merge_cells(start_row=3, start_column=4, end_row=3, end_column=total_col)
    ws2.cell(3, 4).value = date_range_str
    ws2.cell(3, 4).font = Font(bold=True, size=18)
    ws2.cell(3, 4).alignment = Alignment(horizontal="center")

    # Row 4: "Pace Report"
    ws2.merge_cells(start_row=4, start_column=4, end_row=4, end_column=total_col)
    ws2.cell(4, 4).value = "Pace Report"
    ws2.cell(4, 4).font = Font(bold=True, size=18)
    ws2.cell(4, 4).alignment = Alignment(horizontal="center")

    # Row 5: labels
    ws2.cell(5, 4).value = "Total Contracted Block"
    ws2.cell(5, 4).font = Font(italic=True, size=11)
    ws2.cell(5, 4).alignment = Alignment(horizontal="center")

    if cutoff_earliest:
        cutoff_cell = ws2.cell(5, last_col_idx)
        cutoff_cell.value = f"Cut off date: {cutoff_earliest}"
        cutoff_cell.font = Font(bold=True, size=10, color="FFFF0000")

    # ── Date header row (row 6) ───────────────────────────────────────────
    date_hdr_fill = PatternFill("solid", fgColor="1A3A5C")
    date_hdr_font = Font(bold=True, size=11, color="FFFFFF")
    day_hdr_fill  = PatternFill("solid", fgColor="2D5986")

    ws2.cell(6, 4).value = "Date:"
    ws2.cell(6, 4).font = Font(bold=True, size=11, color="FFFFFF")
    ws2.cell(6, 4).fill = date_hdr_fill

    for ni, ed in enumerate(event_dates):
        col_idx = start_col + ni
        try:
            d = _date_xlsx.fromisoformat(ed)
            ws2.cell(6, col_idx).value = d
            ws2.cell(6, col_idx).number_format = "d-mmm"
        except Exception:
            ws2.cell(6, col_idx).value = ed
        ws2.cell(6, col_idx).alignment = Alignment(horizontal="center")
        ws2.cell(6, col_idx).font = date_hdr_font
        ws2.cell(6, col_idx).fill = date_hdr_fill
        ws2.cell(6, col_idx).border = Border(left=Side(style='thin', color='FFFFFF'),
                                              right=Side(style='thin', color='FFFFFF'))

    ws2.cell(6, total_col).value = "Total"
    ws2.cell(6, total_col).font = date_hdr_font
    ws2.cell(6, total_col).alignment = Alignment(horizontal="center")
    ws2.cell(6, total_col).fill = date_hdr_fill

    # ── Day-of-week row (row 7) ───────────────────────────────────────────
    ws2.cell(7, 4).value = "Day:"
    ws2.cell(7, 4).font = Font(bold=True, size=11, color="FFFFFF")
    ws2.cell(7, 4).fill = day_hdr_fill

    day_abbrevs = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']
    for ni, ed in enumerate(event_dates):
        col_idx = start_col + ni
        try:
            d = _date_xlsx.fromisoformat(ed)
            ws2.cell(7, col_idx).value = day_abbrevs[d.weekday()]
        except Exception:
            ws2.cell(7, col_idx).value = ""
        ws2.cell(7, col_idx).alignment = Alignment(horizontal="center")
        ws2.cell(7, col_idx).font = Font(bold=True, size=11, color="FFFFFF")
        ws2.cell(7, col_idx).fill = day_hdr_fill
        ws2.cell(7, col_idx).border = Border(left=Side(style='thin', color='FFFFFF'),
                                              right=Side(style='thin', color='FFFFFF'))
    ws2.cell(7, total_col).fill = day_hdr_fill

    # ── Contracted block rows (rows 8 to 7+n_hotels) ─────────────────────
    hotel_row_fill  = PatternFill("solid", fgColor="F8F9FA")
    hotel_row_fill2 = PatternFill("solid", fgColor="FFFFFF")
    thin = Side(style='thin')
    med  = Side(style='medium')
    def cell_border(top=None, bottom=None, left=None, right=None):
        return Border(top=top, bottom=bottom, left=left, right=right)

    for hi, hd in enumerate(hotel_data):
        r = 8 + hi
        c_cfg = hd['config']
        # cols 1 (A) and 2 (B) intentionally left empty; org (col 3/C) also left empty
        # hotel name goes in col 4 (D) — cols A-C deleted later
        ws2.cell(r, 4).value = f"{hd['hotel_name']} (${hd['contracted_rate']:.0f})"
        ws2.cell(r, 4).font = Font(size=11)
        row_fill = hotel_row_fill if hi % 2 == 0 else hotel_row_fill2

        for ni, ed in enumerate(event_dates):
            col_idx = start_col + ni
            rooms = hd['contracted_block'].get(ed, 0)
            ws2.cell(r, col_idx).value = rooms if rooms else None
            ws2.cell(r, col_idx).alignment = Alignment(horizontal="center")
            ws2.cell(r, col_idx).font = Font(size=11)
            ws2.cell(r, col_idx).fill = row_fill
            ws2.cell(r, col_idx).border = cell_border(top=Side(style='hair'), bottom=Side(style='hair'),
                                                       left=Side(style='hair'), right=Side(style='hair'))

        ws2.cell(r, total_col).value = hd['contracted_total']
        ws2.cell(r, total_col).font = Font(bold=True, size=11)
        ws2.cell(r, total_col).alignment = Alignment(horizontal="center")
        ws2.cell(r, total_col).fill = row_fill
        ws2.cell(r, total_col).border = cell_border(left=Side(style='thin'), right=Side(style='thin'),
                                                      top=Side(style='hair'), bottom=Side(style='hair'))

        # Attrition column
        atr = c_cfg['attrition_pct']
        if atr:
            atr_rooms = round(hd['contracted_total'] * atr)
            ws2.cell(r, pct_col).value = f"Attrition {round(atr*100):.0f}% = {atr_rooms}"
        else:
            ws2.cell(r, pct_col).value = "Attrition Waived"
        ws2.cell(r, pct_col).font = Font(size=10)
        ws2.cell(r, pct_col).fill = row_fill

    # ── Total block row (row 8+n_hotels) ─────────────────────────────────
    total_block_row = 8 + n_hotels
    red_font_bold = Font(bold=True, size=11, color="FFFF0000")
    total_blk_fill = PatternFill("solid", fgColor="FCE4EC")  # pale red
    total_blk_border_mid = Border(top=Side(style='medium'), bottom=Side(style='medium'))
    ws2.cell(total_block_row, 4).value = "Total Block"
    ws2.cell(total_block_row, 4).font = red_font_bold
    ws2.cell(total_block_row, 4).alignment = Alignment(horizontal="center")
    ws2.cell(total_block_row, 4).fill = total_blk_fill
    ws2.cell(total_block_row, 4).border = Border(top=Side(style='medium'), bottom=Side(style='medium'),
                                                   left=Side(style='medium'))

    for ni, ed in enumerate(event_dates):
        col_idx = start_col + ni
        night_total = sum(hd['contracted_block'].get(ed, 0) for hd in hotel_data)
        ws2.cell(total_block_row, col_idx).value = night_total if night_total else None
        ws2.cell(total_block_row, col_idx).font = red_font_bold
        ws2.cell(total_block_row, col_idx).alignment = Alignment(horizontal="center")
        ws2.cell(total_block_row, col_idx).fill = total_blk_fill
        ws2.cell(total_block_row, col_idx).border = total_blk_border_mid

    ws2.cell(total_block_row, total_col).value = combined_block
    ws2.cell(total_block_row, total_col).font = red_font_bold
    ws2.cell(total_block_row, total_col).alignment = Alignment(horizontal="center")
    ws2.cell(total_block_row, total_col).fill = total_blk_fill
    ws2.cell(total_block_row, total_col).border = Border(top=Side(style='medium'), bottom=Side(style='medium'),
                                                          right=Side(style='medium'))

    # ── Spacer row ────────────────────────────────────────────────────────
    spacer_row = total_block_row + 1  # = 9 + n_hotels

    # ── Weekly snapshot blocks ────────────────────────────────────────────
    body_start = spacer_row + 1  # = 10 + n_hotels
    yellow_fill  = PatternFill("solid", fgColor="FFFF00")
    blk_hdr_fill = PatternFill("solid", fgColor="D9E1F2")   # slate-blue header
    blk_tot_fill = PatternFill("solid", fgColor="E2EFDA")   # pale green total row
    blk_h1_fill  = PatternFill("solid", fgColor="F2F7FF")   # hotel alt 1
    blk_h2_fill  = PatternFill("solid", fgColor="FFFFFF")   # hotel alt 2
    hair_b = Border(top=Side(style='hair'), bottom=Side(style='hair'),
                    left=Side(style='hair'), right=Side(style='hair'))
    edge_l = Border(top=Side(style='hair'), bottom=Side(style='hair'),
                    left=Side(style='thin'), right=Side(style='hair'))
    edge_r = Border(top=Side(style='hair'), bottom=Side(style='hair'),
                    left=Side(style='hair'), right=Side(style='thin'))

    for i, report_date in enumerate(all_report_dates):
        block_row = body_start + (i * 8)
        is_latest = (report_date == all_report_dates[-1])
        hdr_fill = yellow_fill if is_latest else blk_hdr_fill
        hdr_color = "000000" if is_latest else "FF0000"

        # Row 0 of block: header ─────────────────────────────────────────
        ws2.cell(block_row, 3).value = "Actual Pick Up"
        ws2.cell(block_row, 3).font = Font(bold=True, size=11, color=hdr_color)
        ws2.cell(block_row, 3).fill = hdr_fill

        try:
            ws2.cell(block_row, 4).value = _date_xlsx.fromisoformat(report_date)
            ws2.cell(block_row, 4).number_format = "d-mmm-yy"
        except Exception:
            ws2.cell(block_row, 4).value = report_date
        ws2.cell(block_row, 4).font = Font(size=10, color=hdr_color)
        ws2.cell(block_row, 4).fill = hdr_fill

        ws2.merge_cells(start_row=block_row, start_column=4, end_row=block_row, end_column=total_col)

        wo = weeks_out('2026-07-17', report_date)
        wo_label = f"{wo} weeks" if wo is not None else ""
        ws2.cell(block_row, last_col_idx).value = wo_label
        ws2.cell(block_row, last_col_idx).font = Font(italic=True, size=11, color=hdr_color)
        ws2.cell(block_row, last_col_idx).fill = hdr_fill

        # Row 1 of block: Total row ──────────────────────────────────────
        tr = block_row + 1
        ws2.cell(tr, 4).value = "Total"
        ws2.cell(tr, 4).font = Font(bold=True, italic=True, size=10)
        ws2.cell(tr, 4).alignment = Alignment(horizontal="center")
        ws2.cell(tr, 4).fill = blk_tot_fill
        ws2.cell(tr, 4).border = Border(top=Side(style='thin'), bottom=Side(style='thin'),
                                         left=Side(style='thin'), right=Side(style='hair'))

        for ni, ed in enumerate(event_dates):
            night_total = total_by_night.get(report_date, {}).get(ed, 0)
            col_idx = start_col + ni
            ws2.cell(tr, col_idx).value = night_total if night_total else None
            ws2.cell(tr, col_idx).number_format = "0"
            ws2.cell(tr, col_idx).font = Font(bold=True, size=11)
            ws2.cell(tr, col_idx).alignment = Alignment(horizontal="center")
            ws2.cell(tr, col_idx).fill = blk_tot_fill
            ws2.cell(tr, col_idx).border = Border(top=Side(style='thin'), bottom=Side(style='thin'),
                                                   left=Side(style='hair'), right=Side(style='hair'))

        ws2.cell(tr, total_col).value = total_rooms_agg.get(report_date, 0) or None
        ws2.cell(tr, total_col).number_format = "0"
        ws2.cell(tr, total_col).font = Font(bold=True, size=11)
        ws2.cell(tr, total_col).fill = blk_tot_fill
        ws2.cell(tr, total_col).border = Border(top=Side(style='thin'), bottom=Side(style='thin'),
                                                  left=Side(style='hair'), right=Side(style='thin'))

        pct_val = total_rooms_agg.get(report_date, 0) / combined_block if combined_block else None
        ws2.cell(tr, pct_col).value = pct_val
        ws2.cell(tr, pct_col).number_format = "0%"
        ws2.cell(tr, pct_col).font = Font(bold=True, size=11)
        ws2.cell(tr, pct_col).fill = blk_tot_fill

        # Rows 2-N of block: per-hotel rows ─────────────────────────────
        for hi, hd in enumerate(hotel_data):
            hr = block_row + 2 + hi
            cid = hd['config']['id']
            h_row_fill = blk_h1_fill if hi % 2 == 0 else blk_h2_fill
            ws2.cell(hr, 4).value = hd['hotel_name']
            ws2.cell(hr, 4).font = Font(bold=True, italic=True, size=10)
            ws2.cell(hr, 4).alignment = Alignment(horizontal="center")
            ws2.cell(hr, 4).fill = h_row_fill
            ws2.cell(hr, 4).border = edge_l

            hotel_pbn_rd = hotel_pbn.get(cid, {}).get(report_date, {})
            for ni, ed in enumerate(event_dates):
                rooms = hotel_pbn_rd.get(ed, None)
                col_idx = start_col + ni
                ws2.cell(hr, col_idx).value = rooms if rooms else None
                ws2.cell(hr, col_idx).number_format = "0"
                ws2.cell(hr, col_idx).alignment = Alignment(horizontal="center")
                ws2.cell(hr, col_idx).font = Font(size=11)
                ws2.cell(hr, col_idx).fill = h_row_fill
                ws2.cell(hr, col_idx).border = hair_b

            # Use stored total_rooms for the hotel
            w_entry = next((w for w in hd['weekly'] if w['report_date'] == report_date), None)
            hotel_total = (w_entry['total_rooms'] or 0) if w_entry else 0

            ws2.cell(hr, total_col).value = hotel_total if hotel_total else None
            ws2.cell(hr, total_col).number_format = "0"
            ws2.cell(hr, total_col).font = Font(bold=True, size=11)
            ws2.cell(hr, total_col).fill = h_row_fill
            ws2.cell(hr, total_col).border = edge_r

            h_pct = hotel_total / hd['contracted_total'] if hd['contracted_total'] else None
            ws2.cell(hr, pct_col).value = h_pct
            ws2.cell(hr, pct_col).number_format = "0%"
            ws2.cell(hr, pct_col).font = Font(bold=True, size=11)
            ws2.cell(hr, pct_col).fill = h_row_fill

    # ── Remove empty cols A–B, then set column widths ────────────────────
    ws2.delete_cols(1, 2)  # cols A (booking_id) + B (event_name) → now C→A, D→B, E→C …
    # After deletion: A=org(empty), B=hotel name, C=first night, … total/pct/weeks-label shift -2
    ws2.column_dimensions['A'].width = 4   # empty spacer
    ws2.column_dimensions['B'].width = 26  # hotel name
    for ni in range(n_nights):
        ws2.column_dimensions[get_column_letter(start_col - 2 + ni)].width = 9
    ws2.column_dimensions[get_column_letter(total_col - 2)].width = 8
    ws2.column_dimensions[get_column_letter(pct_col - 2)].width = 14
    ws2.column_dimensions[get_column_letter(last_col_idx - 2)].width = 10

    # ── Row heights Tab 2 ─────────────────────────────────────────────────
    for rn in range(2, 5):
        ws2.row_dimensions[rn].height = 23.75
    for rn in range(5, ws2.max_row + 1):
        ws2.row_dimensions[rn].height = 14.25

    # ── Return file ───────────────────────────────────────────────────────
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    safe_name = _re_fn.sub(r'[^a-zA-Z0-9_]+', '_', event_name)
    return send_file(buf, as_attachment=True,
                     download_name=f"{safe_name}_Pace_Report.xlsx",
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


def _get_current_pickup_events(db, account_filter=None):
    """Shared helper: return current-section events sorted by org → event_start.
    account_filter: None = admin (no filter), [] = no access, [list] = allowed accounts."""
    from datetime import date as _d, timedelta
    today = _d.today()
    future_cutoff = today + timedelta(days=120)
    if account_filter is None:
        af_where, af_params = '', []
    elif account_filter:
        ph = ','.join('?' * len(account_filter))
        af_where = f' AND (p.AccountName IN ({ph}) OR c.organization IN ({ph}))'
        af_params = account_filter + account_filter
    else:
        af_where, af_params = ' AND 1=0', []
    configs = db.execute(f"""
        SELECT c.*,
               p.StartDate AS bk_start, p.EndDate AS bk_end,
               p.EventName AS bk_event, p.AccountName AS bk_org,
               p.Customer  AS bk_hotel
        FROM pickup_config c
        LEFT JOIN ReportPipeline p ON CAST(c.booking_id AS TEXT)=CAST(p.BookingId AS TEXT)
        WHERE c.status='active'{af_where}
    """, af_params).fetchall()
    current_events = []
    for c in configs:
        block            = json.loads(c['contracted_block'] or '{}')
        all_dates        = sorted(block.keys())
        contracted_total = sum(block.values())
        start_date = None
        if all_dates:
            try:
                start_date = _d.fromisoformat(all_dates[0])
            except Exception:
                pass
        if start_date is None and c['bk_start']:
            try:
                _bs = _to_iso(c['bk_start'])
                if _bs:
                    start_date = _d.fromisoformat(_bs)
            except Exception:
                pass
        has_final = db.execute(
            "SELECT COUNT(*) FROM pickup_weekly WHERE config_id=? AND LOWER(label) LIKE '%final%'",
            (c['id'],)
        ).fetchone()[0] > 0
        if has_final or bool(c['force_past']):
            continue
        if not c['force_current']:
            if start_date is not None and start_date < today:
                continue
            if start_date is not None and start_date > future_cutoff:
                continue
        last_w = db.execute(
            """SELECT total_rooms, pct_of_block FROM pickup_weekly
               WHERE config_id=? AND (label IS NULL OR label NOT LIKE '%final%')
               ORDER BY report_date DESC LIMIT 1""",
            (c['id'],)
        ).fetchone()
        weekly_rows = db.execute(
            """SELECT report_date, total_rooms, pct_of_block, pct_of_attrition, pickup_by_night
               FROM pickup_weekly
               WHERE config_id=?
                 AND (label IS NULL
                      OR (label NOT LIKE 'ASAS %'
                          AND label NOT LIKE 'Historical%'
                          AND LOWER(label) NOT LIKE '%final%'))
               ORDER BY report_date DESC""",
            (c['id'],)
        ).fetchall()
        org        = c['organization'] or c['bk_org'] or ''
        event_name = c['event_name']   or c['bk_event'] or org
        end_date   = None
        if all_dates:
            try:
                end_date = _d.fromisoformat(all_dates[-1])
            except Exception:
                pass
        # Extend event_dates to include any shoulder nights recorded in pickup_by_night
        all_event_date_set = set(all_dates)
        for w in weekly_rows:
            try:
                pbn = json.loads(w['pickup_by_night'] or '{}')
                for d in pbn.keys():
                    if d:
                        all_event_date_set.add(d)
            except Exception:
                pass
        all_event_dates = sorted(all_event_date_set)

        current_events.append({
            'config':           c,
            'block':            block,
            'event_dates':      all_event_dates,
            'contracted_total': contracted_total,
            'org':              org,
            'event_name':       event_name,
            'hotel':            c['hotel'] or c['bk_hotel'] or '',
            'start_date':       start_date,
            'end_date':         end_date,
            'weekly':           [dict(w) for w in weekly_rows],
            'last_total':       last_w['total_rooms']  if last_w else 0,
            'last_pct':         last_w['pct_of_block'] if last_w else None,
        })
    current_events.sort(key=lambda r: (r['org'].lower(), str(r['start_date'] or '')))
    return current_events


@app.route('/pickup/customer-report', methods=['GET'])
def pickup_customer_report_select():
    """Selection page — choose which customers/events to include in the report."""
    user = get_current_user()
    db = get_db()
    current_events = _get_current_pickup_events(db, account_filter=get_pickup_account_filter(user))

    # Group by organisation for the UI
    from collections import OrderedDict
    orgs = OrderedDict()
    for ev in current_events:
        key = ev['org'] or '(No Organisation)'
        orgs.setdefault(key, []).append(ev)

    return render_template('pickup_customer_report_select.html',
                           orgs=orgs, total=len(current_events))


@app.route('/pickup/customer-report-xlsx', methods=['GET', 'POST'])
def pickup_customer_report_xlsx():
    """Generate multi-tab XLSX: Tab 1 = summary sorted by customer, Tab N = per-event pace report."""
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from datetime import date as _d
    import io, re as _re

    user  = get_current_user()
    db    = get_db()
    today = _d.today()

    # Load all current events via shared helper (filtered by account access)
    all_events = _get_current_pickup_events(db, account_filter=get_pickup_account_filter(user))

    # If POSTed from the selection page, filter to chosen config IDs only
    if request.method == 'POST':
        selected_ids = set(request.form.getlist('cids'))
        if selected_ids:
            current_events = [ev for ev in all_events
                              if str(ev['config']['id']) in selected_ids]
        else:
            current_events = all_events   # nothing checked → include all
    else:
        current_events = all_events       # direct GET → include all

    if not current_events:
        flash('No events selected — nothing to download.', 'warning')
        return redirect(url_for('pickup_customer_report_select'))

    # ── Colour / style constants ─────────────────────────────────────────────
    navy       = "1A3A5C"
    mid_navy   = "2D5986"
    setup_fg   = "BFBFBF"   # rows 9–12 gray
    green_fg   = "92D050"   # final / remaining rows
    org_hdr_fg = "D9E1F2"   # org separator in summary

    def _fill(hex6):      return PatternFill("solid", fgColor=hex6)
    def _font(bold=False, size=11, color="000000", italic=False):
        return Font(bold=bold, size=size, color=color, italic=italic)

    hair  = Side(style='hair')
    thin  = Side(style='thin')
    def _border(**kw): return Border(**kw)

    wb = Workbook()

    # ════════════════════════════════════════════════════════════════════════
    # TAB 1 — SUMMARY
    # ════════════════════════════════════════════════════════════════════════
    ws1       = wb.active
    ws1.title = "Summary"

    SUMM_COLS = 13
    last_ltr  = get_column_letter(SUMM_COLS)

    # Row 1 – title banner
    ws1.merge_cells(f"A1:{last_ltr}1")
    ws1["A1"].value     = f"PICK-UP SUMMARY REPORT — {today.year}"
    ws1["A1"].font      = Font(bold=True, size=18, color="FFFFFF")
    ws1["A1"].fill      = _fill(navy)
    ws1["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[1].height = 30

    # Row 2 – sub-header
    ws1.merge_cells(f"A2:{last_ltr}2")
    ws1["A2"].value     = (f"Generated: {today.strftime('%B %d, %Y')}   |   "
                           f"Current events sorted by customer")
    ws1["A2"].font      = Font(italic=True, size=10, color="555555")
    ws1["A2"].alignment = Alignment(horizontal="left", vertical="center")
    ws1.row_dimensions[2].height = 16

    # Row 3 – column headers
    hdrs = ["Organization", "Event Name", "Hotel", "Start", "End",
            "Block", "Pickup", "WoW Chg", "% Block", "% Attrition",
            "Cutoff", "Days Left", "Status"]
    for ci, h in enumerate(hdrs, 1):
        cell            = ws1.cell(3, ci, h)
        cell.font       = Font(bold=True, color="FFFFFF", size=11)
        cell.fill       = _fill(mid_navy)
        cell.alignment  = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[3].height = 18

    # Column widths for summary
    for ci, w in enumerate([30,30,22,11,11,10,10,10,9,12,12,9,10], 1):
        ws1.column_dimensions[get_column_letter(ci)].width = w

    # Data rows
    row_idx  = 4
    prev_org = None
    alt_fills = [_fill("F2F7FF"), _fill("FFFFFF")]
    green_txt  = Font(bold=True, size=11, color="006400")
    yellow_txt = Font(bold=True, size=11, color="7D6608")
    red_txt    = Font(bold=True, size=11, color="CC0000")

    for ev in current_events:
        # Org separator row when organisation changes
        if ev['org'] != prev_org:
            ws1.merge_cells(f"A{row_idx}:{last_ltr}{row_idx}")
            cell            = ws1.cell(row_idx, 1, ev['org'] or '—')
            cell.font       = Font(bold=True, size=11, color=navy)
            cell.fill       = _fill(org_hdr_fg)
            cell.alignment  = Alignment(horizontal="left", vertical="center")
            ws1.row_dimensions[row_idx].height = 16
            row_idx += 1
            prev_org = ev['org']

        weekly = ev['weekly']
        last_w = weekly[0] if weekly else None
        last_total     = last_w['total_rooms']     if last_w else 0
        last_pct_block = last_w['pct_of_block']    if last_w else None
        last_pct_attr  = last_w['pct_of_attrition'] if last_w else None
        wow = None
        if len(weekly) >= 2 and last_w:
            wow = (last_w['total_rooms'] or 0) - (weekly[1]['total_rooms'] or 0)
        days_left = None
        if ev['config']['cutoff_date']:
            try:
                days_left = (_d.fromisoformat(ev['config']['cutoff_date']) - today).days
            except Exception:
                pass

        status = ('No data' if last_pct_block is None
                  else 'On Pace' if last_pct_block >= 80
                  else 'Watch'   if last_pct_block >= 60
                  else 'At Risk')

        rf = alt_fills[row_idx % 2]
        base_font = _font(size=11)

        def _s(col, val):
            c = ws1.cell(row_idx, col, val)
            c.font = base_font; c.fill = rf
            return c

        _s(1, ev['org'])
        _s(2, ev['event_name'])
        _s(3, ev['hotel'])
        d4 = _s(4, ev['start_date'])
        d4.number_format = "mm-dd-yy"
        d5 = _s(5, ev['end_date'])
        d5.number_format = "mm-dd-yy"
        _s(6, ev['contracted_total']).number_format = "#,##0"
        _s(7, last_total).number_format = "#,##0"
        if wow is not None:
            _s(8, wow).number_format = "+#,##0;-#,##0;0"
        else:
            _s(8, None)

        # % Block — colour-coded
        pct_cell = ws1.cell(row_idx, 9,
                             (last_pct_block / 100) if last_pct_block is not None else None)
        pct_cell.fill          = rf
        pct_cell.number_format = "0%"
        if last_pct_block is not None:
            pct_cell.font = (green_txt  if last_pct_block >= 80
                             else yellow_txt if last_pct_block >= 60
                             else red_txt)

        # % Attrition
        a_cell = ws1.cell(row_idx, 10,
                          (last_pct_attr / 100) if last_pct_attr is not None else None)
        a_cell.fill = rf
        a_cell.number_format = "0%"
        a_cell.font = base_font

        # Cutoff date
        cutoff_cell = ws1.cell(row_idx, 11)
        cutoff_cell.fill = rf
        if ev['config']['cutoff_date']:
            try:
                cutoff_cell.value = _d.fromisoformat(ev['config']['cutoff_date'])
                cutoff_cell.number_format = "mm-dd-yy"
            except Exception:
                cutoff_cell.value = ev['config']['cutoff_date']
        cutoff_cell.font = base_font

        dl_cell = ws1.cell(row_idx, 12, days_left)
        dl_cell.fill = rf
        dl_cell.font = (Font(bold=True, size=11, color="CC0000")
                        if days_left is not None and days_left < 0
                        else Font(bold=True, size=11, color="B8860B")
                        if days_left is not None and days_left <= 14
                        else base_font)

        st_cell = ws1.cell(row_idx, 13, status)
        st_cell.fill = rf
        st_cell.font = (green_txt  if status == 'On Pace'
                        else yellow_txt if status == 'Watch'
                        else red_txt    if status == 'At Risk'
                        else base_font)

        ws1.row_dimensions[row_idx].height = 15
        row_idx += 1

    # ════════════════════════════════════════════════════════════════════════
    # TABS 2+ — PER-EVENT PICK-UP UPDATE (matches example format)
    # ════════════════════════════════════════════════════════════════════════
    setup_fill = _fill(setup_fg)
    green_fill = _fill(green_fg)
    used_names = {"Summary"}
    day_abbrevs = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']

    for ev in current_events:
        # ── Tab name ────────────────────────────────────────────────────────
        raw = (ev['event_name'] or ev['org'] or 'Event')
        tab = _re.sub(r'[:\\/?*\[\]]', '', raw).strip()[:31]
        orig_tab = tab
        i = 2
        while tab in used_names:
            tab = orig_tab[:28] + f' {i}'; i += 1
        used_names.add(tab)

        ws = wb.create_sheet(title=tab)
        c  = ev['config']
        event_dates      = ev['event_dates']
        block            = ev['block']
        contracted_total = ev['contracted_total']
        n_nights         = len(event_dates)

        # ── No dates: stub ───────────────────────────────────────────────────
        if n_nights == 0:
            ws["A1"].value = "PICK-UP UPDATE"
            ws["A1"].font  = Font(bold=True, size=18)
            ws["A2"].value = "ORGANIZATION:"
            ws["A2"].font  = Font(bold=True, size=11)
            ws["B2"].value = ev['org']
            ws["A3"].value = "(Contracted block dates not yet defined)"
            ws["A3"].font  = Font(italic=True, size=10, color="888888")
            continue

        # ── Column index helpers ─────────────────────────────────────────────
        S_COL          = 2                            # first night col (B)
        E_COL          = S_COL + n_nights - 1         # last night col
        TOT_COL        = E_COL + 1                    # Total
        CHG_COL        = TOT_COL + 1                  # Change
        PCT_BLK_COL    = CHG_COL + 1                  # % Block
        PCT_ATR_COL    = PCT_BLK_COL + 1              # % Attrition
        LAST_COL       = PCT_ATR_COL
        s_ltr          = get_column_letter(S_COL)
        e_ltr          = get_column_letter(E_COL)
        tot_ltr        = get_column_letter(TOT_COL)
        chg_ltr        = get_column_letter(CHG_COL)
        pct_blk_ltr    = get_column_letter(PCT_BLK_COL)
        pct_atr_ltr    = get_column_letter(PCT_ATR_COL)
        last_ltr_ev    = get_column_letter(LAST_COL)

        # ── Column widths ────────────────────────────────────────────────────
        ws.column_dimensions['A'].width = 24.5
        for ci in range(S_COL, LAST_COL + 1):
            ws.column_dimensions[get_column_letter(ci)].width = (
                10 if ci == PCT_ATR_COL else 9)

        # ── Row 1: Title ─────────────────────────────────────────────────────
        ws.merge_cells(f"A1:{last_ltr_ev}1")
        ws["A1"].value     = "PICK-UP UPDATE"
        ws["A1"].font      = Font(bold=True, size=18)
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 28

        # ── Helper: label in A, merged value in B–last ──────────────────────
        def _hdr(row, label, val):
            ws.cell(row, 1, label).font      = Font(bold=True, size=11)
            ws.cell(row, 1).alignment        = Alignment(horizontal="right", vertical="center")
            ws.merge_cells(f"B{row}:{last_ltr_ev}{row}")
            ws.cell(row, 2, val).font        = Font(size=11)
            ws.cell(row, 2).alignment        = Alignment(horizontal="left", wrap_text=True)

        _hdr(2, "ORGANIZATION:", ev['org'])
        _hdr(3, "HOTEL & LOCATION:", ev['hotel'])

        # Event name + date range
        if event_dates:
            try:
                ds = _d.fromisoformat(event_dates[0]).strftime('%B %d')
                de = _d.fromisoformat(event_dates[-1]).strftime('%B %d, %Y')
                ev_date_str = f"{ds} – {de}"
            except Exception:
                ev_date_str = f"{event_dates[0]} – {event_dates[-1]}"
        else:
            ev_date_str = ""
        _hdr(4, "NAME & DATE OF EVENT:", f"{ev['event_name']}   |   {ev_date_str}")

        # Row 5: Group contact
        ws.cell(5, 1, "GROUP CONTACT").font      = Font(bold=True, size=11)
        ws.cell(5, 1).alignment                  = Alignment(horizontal="right")
        ws.cell(5, 2, c['group_contact'] or '').font = Font(size=11)
        if c['group_contact_email']:
            ec = ws.cell(5, min(4, LAST_COL), c['group_contact_email'])
            ec.font = Font(size=10, color="0070C0")

        _hdr(6, "HOTEL CONTACT:", c['hotel_contact'] or '')
        _hdr(7, "CONTACT'S NUMBER:", c['hotel_contact_email'] or '')
        ws.cell(8, 1, "Booking").font      = Font(bold=True, size=11)
        ws.cell(8, 1).alignment            = Alignment(horizontal="right")
        ws.cell(8, 2, c['booking_id']).font = Font(size=11)

        # ── Row 9: Cut-off date + attrition (gray) ───────────────────────────
        ws.row_dimensions[9].height = 45

        def _setup(row):
            for ci in range(1, LAST_COL + 1):
                ws.cell(row, ci).fill = setup_fill

        _setup(9)
        ws.cell(9, 1, "Cut-off Date:").font      = Font(bold=True, size=11)
        ws.cell(9, 1).alignment                  = Alignment(horizontal="right", vertical="center")
        merge_end = min(S_COL + 2, E_COL)
        ws.merge_cells(start_row=9, start_column=S_COL, end_row=9, end_column=merge_end)
        if c['cutoff_date']:
            try:
                ws.cell(9, S_COL).value = _d.fromisoformat(c['cutoff_date'])
                ws.cell(9, S_COL).number_format = "[$-409]mmmm d, yyyy"
            except Exception:
                ws.cell(9, S_COL).value = c['cutoff_date']
        ws.cell(9, S_COL).font      = Font(size=11)
        ws.cell(9, S_COL).alignment = Alignment(horizontal="left", vertical="center")

        attrition_pct   = c['attrition_pct'] or 0
        attrition_rooms = round(contracted_total * attrition_pct) if contracted_total and attrition_pct else 0

        ws.cell(9, TOT_COL, "Attrition:").font      = Font(bold=True, size=11)
        ws.cell(9, CHG_COL).value                   = attrition_pct or None
        ws.cell(9, CHG_COL).number_format            = "0%"
        ws.cell(9, CHG_COL).font                     = Font(size=11)
        ws.cell(9, PCT_BLK_COL).value               = attrition_rooms or None
        ws.cell(9, PCT_BLK_COL).number_format        = "#,##0"
        ws.cell(9, PCT_BLK_COL).font                 = Font(size=11)
        if c['contracted_rate']:
            ws.cell(9, PCT_ATR_COL).value            = c['contracted_rate']
            ws.cell(9, PCT_ATR_COL).number_format    = '"$"#,##0.00'
            ws.cell(9, PCT_ATR_COL).font             = Font(size=11)

        # ── Row 10: Dates (gray) ─────────────────────────────────────────────
        _setup(10)
        ws.cell(10, 1, "Dates:").font      = Font(bold=True, size=11)
        ws.cell(10, 1).alignment           = Alignment(horizontal="right")
        for ni, ed in enumerate(event_dates):
            ci = S_COL + ni
            try:
                ws.cell(10, ci).value = _d.fromisoformat(ed)
                ws.cell(10, ci).number_format = "mm-dd-yy"
            except Exception:
                ws.cell(10, ci).value = ed
            ws.cell(10, ci).alignment = Alignment(horizontal="center")
            ws.cell(10, ci).font      = Font(size=11)
        ws.cell(10, TOT_COL, "Total").font      = Font(bold=True, size=11)
        ws.cell(10, TOT_COL).alignment          = Alignment(horizontal="center")

        # ── Row 11: Day of week (gray) ───────────────────────────────────────
        _setup(11)
        ws.cell(11, 1, "Day:").font    = Font(bold=True, size=11)
        ws.cell(11, 1).alignment       = Alignment(horizontal="right")
        for ni, ed in enumerate(event_dates):
            ci = S_COL + ni
            try:
                ws.cell(11, ci).value = day_abbrevs[_d.fromisoformat(ed).weekday()]
            except Exception:
                pass
            ws.cell(11, ci).alignment = Alignment(horizontal="center")
            ws.cell(11, ci).font      = Font(size=11)

        # ── Row 12: Block (gray) ─────────────────────────────────────────────
        _setup(12)
        ws.row_dimensions[12].height = 18
        ws.cell(12, 1, "Block:").font      = Font(bold=True, size=11)
        ws.cell(12, 1).alignment           = Alignment(horizontal="right")
        for ni, ed in enumerate(event_dates):
            ci = S_COL + ni
            ws.cell(12, ci).value         = block.get(ed, 0) or None
            ws.cell(12, ci).number_format = "#,##0_);[Red](#,##0)"
            ws.cell(12, ci).alignment     = Alignment(horizontal="center")
            ws.cell(12, ci).font          = Font(size=11)
        ws.cell(12, TOT_COL).value         = f"=SUM({s_ltr}12:{e_ltr}12)"
        ws.cell(12, TOT_COL).number_format = "#,##0_);[Red](#,##0)"
        ws.cell(12, TOT_COL).font          = Font(bold=True, size=11)
        ws.cell(12, TOT_COL).alignment     = Alignment(horizontal="center")
        for lbl, ci in [("Change", CHG_COL), ("% Block", PCT_BLK_COL), ("% Attrition", PCT_ATR_COL)]:
            ws.cell(12, ci, lbl).font      = Font(bold=True, size=11)
            ws.cell(12, ci).alignment      = Alignment(horizontal="center")

        # ── Row 13: Pending History (green) ──────────────────────────────────
        ws.row_dimensions[13].height = 18
        ws.cell(13, 1, "Pending History").font      = Font(bold=True, size=11)
        ws.cell(13, 1).fill                         = green_fill
        ws.cell(13, 1).alignment                    = Alignment(horizontal="left")
        for ci in range(S_COL, LAST_COL + 1):
            ws.cell(13, ci).fill = green_fill

        # ── Rows 14+: Weekly pickup entries (most recent first) ───────────────
        weekly     = ev['weekly']
        n_data     = len(weekly)

        for wi, w in enumerate(weekly):
            rn = 14 + wi
            ws.row_dimensions[rn].height = 15

            # Date label
            try:
                ws.cell(rn, 1).value = _d.fromisoformat(w['report_date'])
                ws.cell(rn, 1).number_format = "mm-dd-yy"
            except Exception:
                ws.cell(rn, 1).value = w['report_date']
            ws.cell(rn, 1).font      = Font(size=11)
            ws.cell(rn, 1).alignment = Alignment(horizontal="left")

            # Per-night pickup
            pbn = {}
            try:
                pbn = json.loads(w['pickup_by_night'] or '{}')
            except Exception:
                pass
            for ni, ed in enumerate(event_dates):
                ci    = S_COL + ni
                rooms = pbn.get(ed)
                ws.cell(rn, ci).value         = rooms if rooms else None
                ws.cell(rn, ci).number_format = "#,##0_);[Red](#,##0)"
                ws.cell(rn, ci).alignment     = Alignment(horizontal="center")
                ws.cell(rn, ci).font          = Font(size=11)

            # Total (SUM formula)
            ws.cell(rn, TOT_COL).value         = f"=SUM({s_ltr}{rn}:{e_ltr}{rn})"
            ws.cell(rn, TOT_COL).number_format = "#,##0_);[Red](#,##0)"
            ws.cell(rn, TOT_COL).font          = Font(bold=True, size=11)
            ws.cell(rn, TOT_COL).alignment     = Alignment(horizontal="center")

            # Change vs next (older) row
            if wi < n_data - 1:
                ws.cell(rn, CHG_COL).value = f"={tot_ltr}{rn}-{tot_ltr}{rn+1}"
            else:
                ws.cell(rn, CHG_COL).value = f"={tot_ltr}{rn}"
            ws.cell(rn, CHG_COL).number_format = "#,##0_);[Red](#,##0)"
            ws.cell(rn, CHG_COL).font          = Font(size=11)
            ws.cell(rn, CHG_COL).alignment     = Alignment(horizontal="center")

            # % of contracted block
            ws.cell(rn, PCT_BLK_COL).value         = f"={tot_ltr}{rn}/{tot_ltr}$12"
            ws.cell(rn, PCT_BLK_COL).number_format = "0%"
            ws.cell(rn, PCT_BLK_COL).font          = Font(size=11)
            ws.cell(rn, PCT_BLK_COL).alignment     = Alignment(horizontal="center")

            # % of attrition requirement
            if attrition_rooms:
                ws.cell(rn, PCT_ATR_COL).value         = f"={tot_ltr}{rn}/{pct_blk_ltr}$9"
                ws.cell(rn, PCT_ATR_COL).number_format = "0%"
                ws.cell(rn, PCT_ATR_COL).font          = Font(size=11)
                ws.cell(rn, PCT_ATR_COL).alignment     = Alignment(horizontal="center")

        # ── Last row: Remaining (green) ───────────────────────────────────────
        rem_row = 14 + n_data
        ws.row_dimensions[rem_row].height = 18
        ws.cell(rem_row, 1, "Remaining").font      = Font(bold=True, size=11)
        ws.cell(rem_row, 1).fill                   = green_fill
        ws.cell(rem_row, 1).alignment              = Alignment(horizontal="left")

        latest_data_row = 14  # most recent weekly entry
        for ni, ed in enumerate(event_dates):
            ci      = S_COL + ni
            col_ltr = get_column_letter(ci)
            if n_data > 0:
                ws.cell(rem_row, ci).value = f"={col_ltr}12-{col_ltr}{latest_data_row}"
            else:
                ws.cell(rem_row, ci).value = f"={col_ltr}12"
            ws.cell(rem_row, ci).number_format = "#,##0_);[Red](#,##0)"
            ws.cell(rem_row, ci).alignment     = Alignment(horizontal="center")
            ws.cell(rem_row, ci).font          = Font(size=11)
            ws.cell(rem_row, ci).fill          = green_fill

        # Remaining total
        ws.cell(rem_row, TOT_COL).value = (
            f"={tot_ltr}12-{tot_ltr}{latest_data_row}" if n_data > 0
            else f"={tot_ltr}12")
        ws.cell(rem_row, TOT_COL).number_format = "#,##0_);[Red](#,##0)"
        ws.cell(rem_row, TOT_COL).font          = Font(bold=True, size=11)
        ws.cell(rem_row, TOT_COL).alignment     = Alignment(horizontal="center")
        ws.cell(rem_row, TOT_COL).fill          = green_fill

    # ── Return file ───────────────────────────────────────────────────────────
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    fn = f"Pickup_Customer_Report_{today.strftime('%Y%m%d')}.xlsx"
    return send_file(buf, as_attachment=True, download_name=fn,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/pickup/<int:cid>/upload-contract', methods=['GET', 'POST'])
def pickup_upload_contract(cid):
    """Upload a PDF/Word contract, extract room block with AI, review then save."""
    from pickup_utils import parse_contract_document
    db = get_db()
    config = db.execute("SELECT * FROM pickup_config WHERE id=?", (cid,)).fetchone()
    if not config:
        flash('Event not found.', 'error')
        return redirect(url_for('pickup_dashboard'))

    if request.method == 'POST':
        action = request.form.get('action', 'upload')

        # ── Step 2: user confirmed the extracted data ──────────────────────
        if action == 'confirm':
            try:
                raw_block = request.form.get('contracted_block', '{}')
                block     = json.loads(raw_block)
                rate_str  = request.form.get('contracted_rate', '').strip()
                cutoff    = request.form.get('cutoff_date', '').strip() or None
                atr_str   = request.form.get('attrition_pct', '').strip()
                rate      = float(rate_str) if rate_str else config['contracted_rate']
                atr       = float(atr_str)  if atr_str  else config['attrition_pct']
                hotel_contact       = request.form.get('hotel_contact', '').strip() or None
                hotel_contact_email = request.form.get('hotel_contact_email', '').strip() or None
                # Secondary hotel contact — defaults to same as primary (the sales person)
                hotel_contact2      = request.form.get('hotel_contact2', '').strip() or hotel_contact
                hotel_contact2_email= request.form.get('hotel_contact2_email', '').strip() or hotel_contact_email
                group_contact       = request.form.get('group_contact', '').strip() or None
                group_contact_email = request.form.get('group_contact_email', '').strip() or None

                contract_data     = request.form.get('_contract_data_b64', '')
                contract_filename = request.form.get('_contract_filename', '')
                import base64
                file_blob = base64.b64decode(contract_data) if contract_data else None

                db.execute('''
                    UPDATE pickup_config
                    SET contracted_block     = ?,
                        contracted_rate      = ?,
                        cutoff_date          = ?,
                        attrition_pct        = ?,
                        block_is_estimated   = 0,
                        contract_filename    = ?,
                        contract_data        = ?,
                        hotel_contact        = COALESCE(?, hotel_contact),
                        hotel_contact_email  = COALESCE(?, hotel_contact_email),
                        hotel_contact2       = COALESCE(?, hotel_contact2),
                        hotel_contact2_email = COALESCE(?, hotel_contact2_email),
                        group_contact        = COALESCE(?, group_contact),
                        group_contact_email  = COALESCE(?, group_contact_email)
                    WHERE id = ?
                ''', (json.dumps(block), rate, cutoff, atr,
                      contract_filename or None, file_blob,
                      hotel_contact, hotel_contact_email,
                      hotel_contact2, hotel_contact2_email,
                      group_contact, group_contact_email, cid))
                db.commit()
                flash('Contract data saved — room block updated.', 'success')
                return redirect(url_for('pickup_event', cid=cid))
            except Exception as e:
                flash(f'Error saving contract data: {e}', 'error')
                return redirect(url_for('pickup_event', cid=cid))

        # ── Step 1: parse the uploaded file ───────────────────────────────
        f = request.files.get('contract_file')
        if not f or not f.filename:
            flash('No file selected.', 'error')
            return redirect(url_for('pickup_upload_contract', cid=cid))

        file_bytes = f.read()
        extracted  = parse_contract_document(file_bytes, filename=f.filename)

        if extracted.get('error'):
            flash(f'Could not parse contract: {extracted["error"]}', 'error')
            return redirect(url_for('pickup_event', cid=cid))

        # Don't round-trip the PDF through the form — file_b64 stays empty
        return render_template('pickup_contract_review.html',
                               config=config, extracted=extracted,
                               filename=f.filename, file_b64='')

    return render_template('pickup_contract_upload.html', config=config)


@app.route('/pickup/<int:cid>/contract/download')
def pickup_contract_download(cid):
    """Download the stored contract file."""
    db = get_db()
    row = db.execute(
        "SELECT contract_filename, contract_data FROM pickup_config WHERE id=?", (cid,)
    ).fetchone()
    if not row or not row['contract_data']:
        flash('No contract on file.', 'error')
        return redirect(url_for('pickup_event', cid=cid))
    buf = io.BytesIO(row['contract_data'])
    return send_file(buf, as_attachment=True,
                     download_name=row['contract_filename'] or f'contract_{cid}.pdf')


def _build_housing_form_wb(config, pipeline):
    """Return a filled openpyxl Workbook for the Housing History Form."""
    from openpyxl import load_workbook
    from openpyxl.styles import Alignment
    from datetime import datetime as _dt

    block        = json.loads(config['contracted_block'] or '{}')
    sorted_dates = sorted(block.keys())
    n_dates      = len(sorted_dates)

    comm_pct   = float(pipeline['CommissionPercent']) if pipeline and pipeline['CommissionPercent'] else 0.0
    booking_id = config['booking_id'] or ''
    org_name   = (pipeline['AccountName'] or '') if pipeline else (config['organization'] or '')
    hotel_name = config['hotel'] or ''
    event_name = config['event_name'] or config['organization'] or ''

    if sorted_dates:
        start_str  = _dt.strptime(sorted_dates[0],  '%Y-%m-%d').strftime('%m/%d/%Y')
        end_str    = _dt.strptime(sorted_dates[-1], '%Y-%m-%d').strftime('%m/%d/%Y')
        event_cell = f"{event_name}  {start_str} – {end_str}"
    else:
        event_cell = event_name

    tmpl_path = os.path.join(os.path.dirname(__file__), 'static', 'housing_history_template.xlsx')
    wb = load_workbook(tmpl_path)
    ws = wb.active

    n_extra   = max(0, n_dates - 10)
    if n_extra > 0:
        ws.insert_cols(13, n_extra)

    total_col = 13 + n_extra
    rate_col  = 14 + n_extra

    ws['J1'] = comm_pct
    ws.cell(row=1, column=13 + n_extra).value = booking_id
    ws.cell(row=1, column=16 + n_extra).value = 'USD'

    ws['C2'] = org_name
    ws['C3'] = hotel_name
    ws['C4'] = event_cell

    days_abbr = ['Mon','Tue','Wed','Thu','Fri','Sat','Sun']
    center    = Alignment(horizontal='center', vertical='center')
    col_start = 3

    for col_offset in range(n_dates):
        col = col_start + col_offset
        for row in (5, 6, 7):
            ws.cell(row=row, column=col).value = None

    for i, d in enumerate(sorted_dates):
        col      = col_start + i
        date_obj = _dt.strptime(d, '%Y-%m-%d')
        c5 = ws.cell(row=5, column=col)
        c6 = ws.cell(row=6, column=col)
        c7 = ws.cell(row=7, column=col)
        c5.value = date_obj;  c5.alignment = center
        c6.value = days_abbr[date_obj.weekday()]; c6.alignment = center
        c7.value = block.get(d, 0); c7.alignment = center

    ws.cell(row=7, column=total_col).value = sum(block.get(d, 0) for d in sorted_dates)
    ws.cell(row=7, column=rate_col).value  = config['contracted_rate'] or 0

    return wb, event_name


@app.route('/pickup/<int:cid>/housing-form')
def pickup_housing_form(cid):
    """Generate a pre-filled Housing History Form Excel and return as download."""
    db = get_db()
    config = db.execute("SELECT * FROM pickup_config WHERE id=?", (cid,)).fetchone()
    if not config:
        flash('Event not found.', 'error')
        return redirect(url_for('pickup_dashboard'))

    pipeline = None
    if config['booking_id']:
        pipeline = db.execute(
            'SELECT * FROM ReportPipeline WHERE BookingId = ?', (config['booking_id'],)
        ).fetchone()

    wb, event_name = _build_housing_form_wb(config, pipeline)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    safe_name = event_name.replace('/', '-').replace(' ', '_')[:50]
    return send_file(buf, as_attachment=True,
                     download_name=f"Housing_History_{safe_name}.xlsx",
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/pickup/<int:cid>/final-history', methods=['GET', 'POST'])
def pickup_final_history(cid):
    """Enter or update the Final History (actual final pickup numbers) for a past event."""
    db = get_db()
    config = db.execute("SELECT * FROM pickup_config WHERE id=?", (cid,)).fetchone()
    if not config:
        flash('Event not found.', 'error')
        return redirect(url_for('pickup_dashboard'))

    last = db.execute(
        """SELECT * FROM pickup_weekly WHERE config_id=?
           AND (total_rooms IS NOT NULL AND total_rooms > 0)
           AND (label IS NULL OR (label NOT LIKE '%Pending%' AND label NOT LIKE '%placeholder%'))
           ORDER BY report_date DESC LIMIT 1""",
        (cid,)
    ).fetchone()

    existing_fh = db.execute(
        "SELECT * FROM pickup_weekly WHERE config_id=? AND label='Final History' ORDER BY id DESC LIMIT 1",
        (cid,)
    ).fetchone()

    if request.method == 'POST':
        f = request.form
        block = json.loads(config['contracted_block'] or '{}')
        pickup_by_night = {}
        for d in block:
            val = f.get(f'night_{d}', '').strip()
            if val != '':
                pickup_by_night[d] = int(val)
        total_rooms      = sum(pickup_by_night.values())
        contracted_total = sum(block.values())
        pct_of_block     = round(total_rooms / contracted_total * 100, 1) if contracted_total else None
        attrition_floor  = contracted_total * (config['attrition_pct'] or 0)
        pct_of_attrition = round(total_rooms / attrition_floor * 100, 1) if attrition_floor else None
        prev = db.execute(
            "SELECT total_rooms FROM pickup_weekly WHERE config_id=? AND label IS NULL "
            "ORDER BY report_date DESC, id DESC LIMIT 1", (cid,)
        ).fetchone()
        change_from_last = (total_rooms - prev['total_rooms']) if prev and prev['total_rooms'] else None
        report_date = f.get('report_date') or datetime.now().strftime('%Y-%m-%d')
        notes = f.get('notes', '').strip() or None

        if existing_fh:
            db.execute('''UPDATE pickup_weekly
                SET report_date=?, pickup_by_night=?, total_rooms=?,
                    change_from_last=?, pct_of_block=?, pct_of_attrition=?, notes=?
                WHERE id=?''',
                (report_date, json.dumps(pickup_by_night), total_rooms,
                 change_from_last, pct_of_block, pct_of_attrition, notes, existing_fh['id']))
        else:
            db.execute('''INSERT INTO pickup_weekly
                (config_id, report_date, pickup_by_night, total_rooms,
                 change_from_last, pct_of_block, pct_of_attrition, label, notes)
                VALUES (?,?,?,?,?,?,?,?,?)''',
                (cid, report_date, json.dumps(pickup_by_night), total_rooms,
                 change_from_last, pct_of_block, pct_of_attrition, 'Final History', notes))
        db.commit()
        flash('Final History saved — event moved to Past Events.', 'success')
        return redirect(url_for('pickup_event', cid=cid))

    block = json.loads(config['contracted_block'] or '{}')
    prefill_entry = existing_fh or last
    last_pickup = json.loads(prefill_entry['pickup_by_night']) if prefill_entry else {}
    return render_template('pickup_final_history_form.html',
                           config=config, block=block,
                           last_pickup=last_pickup, existing_fh=existing_fh,
                           today=datetime.now().strftime('%Y-%m-%d'))


@app.route('/pickup/<int:cid>/config/edit', methods=['GET', 'POST'])
def pickup_edit_event(cid):
    db = get_db()
    config = db.execute("SELECT * FROM pickup_config WHERE id=?", (cid,)).fetchone()
    if not config:
        flash('Event not found.', 'error')
        return redirect(url_for('pickup_dashboard'))
    if request.method == 'POST':
        f = request.form
        contracted_block = {}
        for d, r in zip(f.getlist('block_date'), f.getlist('block_rooms')):
            if d and r:
                contracted_block[d] = int(r)
        attrition_raw = f.get('attrition_pct', '')
        attrition = float(attrition_raw) / 100 if attrition_raw else None
        ota_url = f.get('ota_url', '').strip() or None
        hc_name   = f.get('hotel_contact', '').strip() or None
        hc_email  = f.get('hotel_contact_email', '').strip() or None
        hc2_name  = f.get('hotel_contact2', '').strip() or None
        hc2_email = f.get('hotel_contact2_email', '').strip() or None
        gc_name  = f.get('group_contact', '').strip() or None
        gc_email = f.get('group_contact_email', '').strip() or None
        hotel_contacts = [{'name': hc_name or '', 'email': hc_email or ''}] if (hc_name or hc_email) else []
        # Additional CC recipients
        cc_emails = []
        for n, e in zip(f.getlist('cc_name[]'), f.getlist('cc_email[]')):
            if n.strip() or e.strip():
                cc_emails.append({'name': n.strip(), 'email': e.strip()})
        rooming_list_required = 1 if f.get('rooming_list_required') else 0
        _changes = [
            ('organization',          config['organization'],          f['organization']),
            ('event_name',            config['event_name'],            f.get('event_name')),
            ('hotel',                 config['hotel'],                 f.get('hotel')),
            ('cutoff_date',           config['cutoff_date'],           f.get('cutoff_date')),
            ('contracted_rate',       config['contracted_rate'],       f.get('contracted_rate')),
            ('attrition_pct',         config['attrition_pct'],        attrition),
            ('contracted_block',      config['contracted_block'],      json.dumps(contracted_block)),
            ('hotel_contact',         config['hotel_contact'],        hc_name),
            ('hotel_contact_email',   config['hotel_contact_email'],  hc_email),
            ('hotel_contact2',        config['hotel_contact2'],       hc2_name),
            ('hotel_contact2_email',  config['hotel_contact2_email'], hc2_email),
            ('group_contact',         config['group_contact'],        gc_name),
            ('group_contact_email',   config['group_contact_email'],  gc_email),
            ('notes',                 config['notes'],                f.get('notes')),
            ('rooming_list_required', config['rooming_list_required'], rooming_list_required),
        ]
        db.execute('''
            UPDATE pickup_config SET
            booking_id=?, tab_name=?, organization=?, event_name=?, hotel=?,
            hotel_contact=?, hotel_contact_email=?, hotel_contact2=?, hotel_contact2_email=?,
            hotel_contacts=?, group_contact=?, group_contact_email=?, cutoff_date=?, attrition_pct=?,
            contracted_block=?, contracted_rate=?, shoulder_pre=?,
            shoulder_post=?, hotel_booking_link=?, notes=?, ota_url=?, cc_emails=?,
            event_start=?, event_end=?, rooming_list_required=?
            WHERE id=?
        ''', (
            f.get('booking_id'), f.get('tab_name'), f['organization'],
            f.get('event_name'), f.get('hotel'), hc_name,
            hc_email, hc2_name, hc2_email, json.dumps(hotel_contacts), gc_name,
            gc_email, f.get('cutoff_date'),
            attrition, json.dumps(contracted_block),
            float(f['contracted_rate']) if f.get('contracted_rate') else None,
            int(f.get('shoulder_pre', 3)), int(f.get('shoulder_post', 3)),
            f.get('hotel_booking_link'), f.get('notes'), ota_url, json.dumps(cc_emails),
            f.get('event_start') or None, f.get('event_end') or None,
            rooming_list_required, cid
        ))
        _log_change(db, cid, 'edit_event', _changes)
        db.commit()
        _upsert_contacts(db, hotel_contacts, f.get('hotel', ''), cc_emails, f['organization'])
        db.commit()
        flash('Event updated.', 'success')
        return redirect(url_for('pickup_event', cid=cid))
    pipeline_rate = pipeline_org = pipeline_event = None
    if config['booking_id']:
        row = db.execute(
            'SELECT RoomRate, AccountName, EventName FROM ReportPipeline WHERE CAST(BookingId AS INTEGER) = CAST(? AS INTEGER) LIMIT 1',
            (config['booking_id'],)
        ).fetchone()
        if row:
            if row['RoomRate']:
                pipeline_rate = round(float(row['RoomRate']), 2)
            pipeline_org   = row['AccountName'] or None
            pipeline_event = row['EventName'] or None
    return render_template('pickup_config_form.html', config=config,
                           action=url_for('pickup_edit_event', cid=cid),
                           cancel_url=url_for('pickup_event', cid=cid),
                           pipeline_rate=pipeline_rate,
                           pipeline_org=pipeline_org,
                           pipeline_event=pipeline_event)


@app.route('/pickup/<int:cid>/weekly/new', methods=['GET', 'POST'])
def pickup_weekly_new(cid):
    db = get_db()
    config = db.execute("SELECT * FROM pickup_config WHERE id=?", (cid,)).fetchone()
    if not config:
        flash('Event not found.', 'error')
        return redirect(url_for('pickup_dashboard'))
    last = db.execute(
        "SELECT * FROM pickup_weekly WHERE config_id=? ORDER BY report_date DESC LIMIT 1", (cid,)
    ).fetchone()
    if request.method == 'POST':
        f = request.form
        block = json.loads(config['contracted_block'] or '{}')
        pickup_by_night = {}
        # Collect contracted block nights AND any shoulder nights added via the form
        for key in f:
            if key.startswith('night_'):
                val = f.get(key, '').strip()
                if val != '':
                    d = key[6:]  # strip 'night_'
                    try:
                        pickup_by_night[d] = int(val)
                    except ValueError:
                        pass
        total_rooms = sum(pickup_by_night.values())
        contracted_total = sum(block.values())
        change_from_last = (total_rooms - (last['total_rooms'] or 0)) if last else None
        pct_of_block     = round(total_rooms / contracted_total * 100, 1) if contracted_total else None
        attrition_floor  = contracted_total * (config['attrition_pct'] or 0)
        pct_of_attrition = round(total_rooms / attrition_floor * 100, 1) if attrition_floor else None
        ota_rate = float(f['ota_rate']) if f.get('ota_rate') else None
        weekly_ota_url = f.get('weekly_ota_url', '').strip() or config['ota_url'] or None
        db.execute('''
            INSERT INTO pickup_weekly
            (config_id, report_date, pickup_by_night, total_rooms, change_from_last,
             pct_of_block, pct_of_attrition, ota_rate, ota_url, label, notes)
            VALUES (?,?,?,?,?,?,?,?,?,?,?)
        ''', (cid, f['report_date'], json.dumps(pickup_by_night), total_rooms,
              change_from_last, pct_of_block, pct_of_attrition, ota_rate, weekly_ota_url,
              f.get('label'), f.get('notes')))
        _log_change(db, cid, 'add_report', [
            ('report_date', None, f['report_date']),
            ('total_rooms', last['total_rooms'] if last else None, total_rooms),
            ('pct_of_block', last['pct_of_block'] if last else None, pct_of_block),
        ])
        db.commit()
        flash('Weekly entry saved.', 'success')
        return redirect(url_for('pickup_event', cid=cid))
    block = json.loads(config['contracted_block'] or '{}')
    last_pickup = json.loads(last['pickup_by_night']) if last else {}
    today_str = datetime.today().strftime('%Y-%m-%d')
    # Shoulder nights from last entry that aren't in the contracted block
    extra_dates = sorted(set(last_pickup.keys()) - set(block.keys()))
    return render_template('pickup_weekly_form.html', config=config, block=block,
                           last=last, last_pickup=last_pickup, today=today_str,
                           is_edit=False, entry=None, entry_pickup={}, wid=None,
                           extra_dates=extra_dates)


@app.route('/pickup/<int:cid>/weekly/<int:wid>/edit', methods=['GET', 'POST'])
def pickup_weekly_edit(cid, wid):
    db = get_db()
    config = db.execute("SELECT * FROM pickup_config WHERE id=?", (cid,)).fetchone()
    entry  = db.execute("SELECT * FROM pickup_weekly WHERE id=? AND config_id=?", (wid, cid)).fetchone()
    if not config or not entry:
        flash('Entry not found.', 'error')
        return redirect(url_for('pickup_event', cid=cid))
    if request.method == 'POST':
        f = request.form
        block = json.loads(config['contracted_block'] or '{}')
        pickup_by_night = {}
        # Collect contracted block nights AND any shoulder nights added via the form
        for key in f:
            if key.startswith('night_'):
                val = f.get(key, '').strip()
                if val != '':
                    d = key[6:]  # strip 'night_'
                    try:
                        pickup_by_night[d] = int(val)
                    except ValueError:
                        pass
        total_rooms = sum(pickup_by_night.values())
        contracted_total = sum(block.values())
        pct_of_block     = round(total_rooms / contracted_total * 100, 1) if contracted_total else None
        attrition_floor  = contracted_total * (config['attrition_pct'] or 0)
        pct_of_attrition = round(total_rooms / attrition_floor * 100, 1) if attrition_floor else None
        ota_rate = float(f['ota_rate']) if f.get('ota_rate') else None
        weekly_ota_url = f.get('weekly_ota_url', '').strip() or config['ota_url'] or None
        db.execute('''
            UPDATE pickup_weekly
            SET report_date=?, pickup_by_night=?, total_rooms=?,
                pct_of_block=?, pct_of_attrition=?, ota_rate=?, ota_url=?, label=?, notes=?
            WHERE id=? AND config_id=?
        ''', (f['report_date'], json.dumps(pickup_by_night), total_rooms,
              pct_of_block, pct_of_attrition, ota_rate, weekly_ota_url,
              f.get('label') or None, f.get('notes') or None, wid, cid))
        # Recompute change_from_last for all entries
        all_w = db.execute(
            "SELECT id, total_rooms, label FROM pickup_weekly WHERE config_id=? ORDER BY report_date ASC, id ASC", (cid,)
        ).fetchall()
        prev_total = None
        for w_row in all_w:
            is_label = bool(w_row['label'] and w_row['label'].strip())
            if is_label or w_row['total_rooms'] is None:
                db.execute("UPDATE pickup_weekly SET change_from_last=NULL WHERE id=?", (w_row['id'],))
            else:
                change = (w_row['total_rooms'] - prev_total) if prev_total is not None else None
                db.execute("UPDATE pickup_weekly SET change_from_last=? WHERE id=?", (change, w_row['id']))
                prev_total = w_row['total_rooms']
        db.commit()
        flash('Weekly entry updated.', 'success')
        return redirect(url_for('pickup_event', cid=cid))
    block = json.loads(config['contracted_block'] or '{}')
    entry_pickup = json.loads(entry['pickup_by_night']) if entry['pickup_by_night'] else {}
    # Shoulder nights saved on this entry that aren't in the contracted block
    extra_dates = sorted(set(entry_pickup.keys()) - set(block.keys()))
    prev_entry = db.execute(
        "SELECT * FROM pickup_weekly WHERE config_id=? AND report_date < ? ORDER BY report_date DESC LIMIT 1",
        (cid, entry['report_date'])
    ).fetchone()
    last_pickup = json.loads(prev_entry['pickup_by_night']) if prev_entry and prev_entry['pickup_by_night'] else {}
    return render_template('pickup_weekly_form.html',
                           config=config, block=block,
                           last=prev_entry, last_pickup=last_pickup,
                           entry_pickup=entry_pickup, entry=entry,
                           today=entry['report_date'], is_edit=True, wid=wid,
                           extra_dates=extra_dates)


@app.route('/pickup/<int:cid>/weekly/<int:wid>/delete', methods=['POST'])
def pickup_weekly_delete(cid, wid):
    db = get_db()
    row = db.execute("SELECT * FROM pickup_weekly WHERE id=? AND config_id=?", (wid, cid)).fetchone()
    if row:
        username = session.get('username') or 'unknown'
        db.execute('''
            INSERT INTO pickup_weekly_deleted
            (original_id, config_id, report_date, pickup_by_night, total_rooms,
             change_from_last, pct_of_block, pct_of_attrition, ota_rate, label, notes, deleted_by)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?)
        ''', (row['id'], row['config_id'], row['report_date'], row['pickup_by_night'],
              row['total_rooms'], row['change_from_last'], row['pct_of_block'],
              row['pct_of_attrition'], row['ota_rate'], row['label'], row['notes'], username))
        db.execute("DELETE FROM pickup_weekly WHERE id=? AND config_id=?", (wid, cid))
        _log_change(db, cid, 'delete_report', [
            ('report_date', row['report_date'], None),
            ('total_rooms', row['total_rooms'], None),
        ])
    db.commit()
    flash('Weekly entry deleted. You can restore it within 90 days from the pickup event page.', 'success')
    return redirect(url_for('pickup_event', cid=cid))


@app.route('/pickup/<int:cid>/weekly/deleted/<int:did>/restore', methods=['POST'])
def pickup_weekly_restore(cid, did):
    db = get_db()
    row = db.execute(
        "SELECT * FROM pickup_weekly_deleted WHERE id=? AND config_id=?", (did, cid)
    ).fetchone()
    if row:
        db.execute('''
            INSERT INTO pickup_weekly
            (config_id, report_date, pickup_by_night, total_rooms, change_from_last,
             pct_of_block, pct_of_attrition, ota_rate, label, notes)
            VALUES (?,?,?,?,?,?,?,?,?,?)
        ''', (row['config_id'], row['report_date'], row['pickup_by_night'], row['total_rooms'],
              row['change_from_last'], row['pct_of_block'], row['pct_of_attrition'],
              row['ota_rate'], row['label'], row['notes']))
        db.execute("DELETE FROM pickup_weekly_deleted WHERE id=?", (did,))
        _log_change(db, cid, 'restore_report', [
            ('report_date', None, row['report_date']),
            ('total_rooms', None, row['total_rooms']),
        ])
        db.commit()
        flash(f'Report for {row["report_date"]} restored.', 'success')
    else:
        flash('Deleted entry not found.', 'error')
    return redirect(url_for('pickup_event', cid=cid))


@app.route('/pickup/<int:cid>/rooming-list', methods=['GET', 'POST'])
def pickup_rooming_upload(cid):
    from pickup_utils import parse_rooming_list_pdf, parse_rooming_list_csv, parse_rooming_list_xls, reconcile, _build_guest_result as _bgr
    db = get_db()
    config = db.execute("SELECT * FROM pickup_config WHERE id=?", (cid,)).fetchone()
    if not config:
        flash('Event not found.', 'error')
        return redirect(url_for('pickup_dashboard'))
    if request.method == 'POST':
        uploaded_files = [f for f in request.files.getlist('rooming_pdf') if f and f.filename]
        labels = request.form.getlist('list_label')
        if not uploaded_files:
            flash('No file uploaded.', 'error')
            return redirect(request.url)
        all_guests, combined_nights, filenames, errors = [], {}, [], []
        any_ai_parsed = False
        for i, f in enumerate(uploaded_files):
            label = labels[i].strip() if i < len(labels) and labels[i].strip() else f.filename
            file_bytes = f.read()
            ext = f.filename.rsplit('.', 1)[-1].lower() if '.' in f.filename else ''
            if ext == 'pdf':
                parsed = parse_rooming_list_pdf(file_bytes)
            elif ext == 'csv':
                parsed = parse_rooming_list_csv(file_bytes)
            elif ext in ('xls', 'xlsx'):
                parsed = parse_rooming_list_xls(file_bytes, filename=f.filename)
            else:
                errors.append(f"{f.filename}: unsupported file type (.{ext}). Use PDF, CSV, XLS, or XLSX.")
                continue
            if parsed.get('error'):
                errors.append(f"{f.filename}: {parsed['error']}")
                continue
            if parsed.get('ai_parsed'):
                any_ai_parsed = True
            if parsed.get('ai_error'):
                errors.append(f"AI parsing failed for {f.filename}: {parsed['ai_error']}")
            filenames.append(f.filename)
            for g in parsed['guests']:
                g['source'] = label
            all_guests.extend(parsed['guests'])
            for date_key, rooms in parsed['nights_by_date'].items():
                combined_nights[date_key] = combined_nights.get(date_key, 0) + rooms
        if errors:
            flash('Parse errors: ' + '; '.join(errors), 'error')
        if not all_guests:
            flash('No guest records could be read. The format may not be recognized.', 'error')
            return redirect(request.url)
        last = db.execute(
            "SELECT * FROM pickup_weekly WHERE config_id=? ORDER BY report_date DESC LIMIT 1", (cid,)
        ).fetchone()
        pickup_by_night = json.loads(last['pickup_by_night']) if last else {}
        recon = reconcile(pickup_by_night, combined_nights)
        combined_result = _bgr(all_guests)
        unique_rooms  = combined_result['unique_rooms']
        combined_nights = combined_result['nights_by_date']
        combined_filename = ', '.join(filenames)
        rl_id = db.execute('''
            INSERT INTO pickup_rooming_list
            (config_id, weekly_id, filename, total_guests,
             nights_by_date, reconciliation_status, discrepancy_notes, guests_json)
            VALUES (?,?,?,?,?,?,?,?)
        ''', (cid, last['id'] if last else None, combined_filename,
              unique_rooms, json.dumps(combined_nights),
              recon['status'], recon.get('notes', ''),
              json.dumps(all_guests))).lastrowid
        db.commit()
        return redirect(url_for('pickup_rooming_review', cid=cid, rl_id=rl_id,
                                ai_parsed=1 if any_ai_parsed else 0))
    return render_template('pickup_rooming_upload.html', config=config)


@app.route('/pickup/<int:cid>/rooming-list/manual', methods=['POST'])
def pickup_rooming_manual(cid):
    from pickup_utils import reconcile
    db = get_db()
    config = db.execute("SELECT * FROM pickup_config WHERE id=?", (cid,)).fetchone()
    if not config:
        flash('Event not found.', 'error')
        return redirect(url_for('pickup_dashboard'))
    dates  = request.form.getlist('manual_date')
    rooms  = request.form.getlist('manual_rooms')
    label  = request.form.get('manual_label', '').strip() or 'Manual entry'
    nights_by_date = {}
    for d, r in zip(dates, rooms):
        if d and r:
            try:
                nights_by_date[d] = int(r)
            except ValueError:
                pass
    if not nights_by_date:
        flash('No room counts entered.', 'error')
        return redirect(url_for('pickup_rooming_upload', cid=cid))
    total_rooms = sum(nights_by_date.values())
    last = db.execute(
        "SELECT * FROM pickup_weekly WHERE config_id=? ORDER BY report_date DESC LIMIT 1", (cid,)
    ).fetchone()
    pickup_by_night = json.loads(last['pickup_by_night']) if last else {}
    recon = reconcile(pickup_by_night, nights_by_date)
    rl_id = db.execute('''
        INSERT INTO pickup_rooming_list
        (config_id, weekly_id, filename, total_guests,
         nights_by_date, reconciliation_status, discrepancy_notes, guests_json)
        VALUES (?,?,?,?,?,?,?,?)
    ''', (cid, last['id'] if last else None, label,
          total_rooms, json.dumps(nights_by_date),
          recon['status'], recon.get('notes', ''), json.dumps([]))).lastrowid
    db.commit()
    return redirect(url_for('pickup_rooming_review', cid=cid, rl_id=rl_id))


@app.route('/pickup/<int:cid>/rooming-list/<int:rl_id>/review')
def pickup_rooming_review(cid, rl_id):
    from pickup_utils import _build_guest_result as _bgr
    from datetime import datetime as _dt, timedelta as _td
    db = get_db()
    config = db.execute("SELECT * FROM pickup_config WHERE id=?", (cid,)).fetchone()
    rl     = db.execute("SELECT * FROM pickup_rooming_list WHERE id=? AND config_id=?", (rl_id, cid)).fetchone()
    last   = db.execute("SELECT * FROM pickup_weekly WHERE config_id=? ORDER BY report_date DESC LIMIT 1", (cid,)).fetchone()
    if not rl:
        flash('Rooming list not found.', 'error')
        return redirect(url_for('pickup_event', cid=cid))
    guests = json.loads(rl['guests_json'] or '[]')
    nights_by_date  = json.loads(rl['nights_by_date'] or '{}')
    pickup_by_night = json.loads(last['pickup_by_night']) if last else {}
    block = json.loads(config['contracted_block'] or '{}')
    dedup_result  = _bgr(guests)
    dedup_summary = dedup_result['dedup_summary']
    unique_rooms  = dedup_result['unique_rooms']
    sources_seen, guests_by_source, nights_by_source = [], {}, {}
    for g in guests:
        src = g.get('source', 'Unknown')
        if src not in sources_seen:
            sources_seen.append(src)
            guests_by_source[src] = []
            nights_by_source[src] = {}
        guests_by_source[src].append(g)
        try:
            arr = _dt.strptime(g['arrival'],   '%Y-%m-%d')
            dep = _dt.strptime(g['departure'], '%Y-%m-%d')
            cur = arr
            while cur < dep:
                key = cur.strftime('%Y-%m-%d')
                nights_by_source[src][key] = nights_by_source[src].get(key, 0) + g.get('rooms', 1)
                cur += _td(days=1)
        except Exception:
            pass
    ai_parsed = bool(int(request.args.get('ai_parsed', 0)))
    return render_template('pickup_rooming_review.html', config=config, rl=rl,
                           guests=guests, nights_by_date=nights_by_date,
                           pickup_by_night=pickup_by_night, block=block,
                           sources=sources_seen, guests_by_source=guests_by_source,
                           nights_by_source=nights_by_source,
                           dedup_summary=dedup_summary, unique_rooms=unique_rooms,
                           ai_parsed=ai_parsed)


@app.route('/pickup/<int:cid>/rooming-list/<int:rl_id>/confirm', methods=['POST'])
def pickup_rooming_confirm(cid, rl_id):
    db = get_db()
    db.execute("UPDATE pickup_rooming_list SET reconciliation_status=? WHERE id=?",
               (request.form.get('status', 'match'), rl_id))
    db.commit()
    flash('Rooming list confirmed.', 'success')
    return redirect(url_for('pickup_event', cid=cid))


@app.route('/pickup/<int:cid>/rooming-list/<int:rl_id>/download-csv')
def pickup_rooming_download_csv(cid, rl_id):
    import csv
    from flask import Response
    db = get_db()
    rl = db.execute("SELECT filename, guests_json FROM pickup_rooming_list WHERE id=? AND config_id=?",
                    (rl_id, cid)).fetchone()
    if not rl:
        flash('Rooming list not found.', 'error')
        return redirect(url_for('pickup_event', cid=cid))
    guests = json.loads(rl['guests_json'] or '[]')
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(['Name', 'Confirmation #', 'Arrival', 'Departure', 'Nights', 'Rooms', 'List'])
    for g in guests:
        writer.writerow([g.get('name',''), g.get('conf_no',''), g.get('arrival',''),
                         g.get('departure',''), g.get('nights',''), g.get('rooms',1), g.get('source','')])
    csv_bytes = output.getvalue().encode('utf-8-sig')
    base = (rl['filename'] or 'rooming_list').rsplit('.', 1)[0]
    return Response(csv_bytes, mimetype='text/csv',
                    headers={'Content-Disposition': f'attachment; filename="{base}_parsed.csv"'})


def _open_in_outlook(cid, subject, to_addr, cc_list, body_html, email_type,
                     attachments=None):
    """Open an Outlook compose window with subject, recipients, HTML body, and optional attachments.
    Returns (True, None) on success or (False, error_str) on failure.
    Only works on macOS (osascript).
    """
    import subprocess, tempfile, os
    from datetime import date as _date

    def esc(s):
        return str(s or '').replace('\\', '\\\\').replace('"', '\\"')

    tmp = tempfile.NamedTemporaryFile(mode='w', suffix='.html', delete=False, encoding='utf-8')
    tmp.write(body_html)
    tmp.close()

    to_line = (
        f'make new to recipient at theMsg with properties '
        f'{{email address:{{address:"{esc(to_addr)}"}}}}'
        if to_addr else ''
    )
    cc_lines = '\n    '.join(
        f'make new cc recipient at theMsg with properties {{email address:{{address:"{esc(a)}"}}}}'
        for a in cc_list if a
    )
    attach_lines = '\n    '.join(
        f'make new attachment at theMsg with properties {{file:(POSIX file "{esc(p)}") as alias}}'
        for p in (attachments or []) if p and os.path.exists(p)
    )

    script = f'''
tell application "Microsoft Outlook"
    set htmlContent to (read (POSIX file "{tmp.name}") as «class utf8»)
    set theMsg to make new outgoing message with properties {{subject:"{esc(subject)}", content:htmlContent}}
    {to_line}
    {cc_lines}
    {attach_lines}
    open theMsg
    activate
end tell
'''
    try:
        subprocess.run(['/usr/bin/osascript', '-e', script], timeout=20)
        attach_note = f' + {len(attachments)} attachment(s)' if attachments else ''
        db = get_db()
        db.execute(
            "INSERT INTO pickup_contact_log (config_id, contact_date, contact_type, notes) VALUES (?,?,?,?)",
            (cid, _date.today().isoformat(), 'email_sent',
             f'Opened in Outlook — {email_type} email{attach_note} — To: {to_addr}')
        )
        db.commit()
        return True, None
    except Exception as e:
        return False, str(e)
    finally:
        try:
            os.unlink(tmp.name)
        except Exception:
            pass


def _create_outlook_draft(user_id, subject, to_addr, cc_list, body_html,
                           attachment_bytes=None, attachment_filename=None):
    """Create a draft in the user's Outlook Drafts folder via Microsoft Graph API.
    Returns (draft_id, None) on success or (None, error_str) on failure.
    """
    import time, base64, requests as _req
    from config import MS_CLIENT_ID, MS_TENANT_ID, MS_CLIENT_SECRET

    db = get_db()
    row = db.execute(
        'SELECT access_token, refresh_token, expires_at FROM UserMicrosoftTokens WHERE user_id = ?',
        (user_id,)
    ).fetchone()
    if not row:
        return None, 'No Microsoft account connected. Visit My Account to connect.'

    access_token  = row['access_token']
    refresh_token = row['refresh_token']
    expires_at    = row['expires_at']

    # Refresh token if expired
    if time.time() >= expires_at:
        r = _req.post(
            f'https://login.microsoftonline.com/{MS_TENANT_ID}/oauth2/v2.0/token',
            data={
                'grant_type':    'refresh_token',
                'client_id':     MS_CLIENT_ID,
                'client_secret': MS_CLIENT_SECRET,
                'refresh_token': refresh_token,
                'scope':         'Mail.ReadWrite User.Read offline_access',
            }
        )
        if r.status_code != 200:
            db.execute('DELETE FROM UserMicrosoftTokens WHERE user_id = ?', (user_id,))
            db.commit()
            return None, 'Microsoft token expired. Please reconnect via My Account.'
        tok           = r.json()
        access_token  = tok['access_token']
        refresh_token = tok.get('refresh_token', refresh_token)
        expires_at    = time.time() + tok.get('expires_in', 3600) - 60
        db.execute(
            'UPDATE UserMicrosoftTokens SET access_token=?, refresh_token=?, expires_at=? WHERE user_id=?',
            (access_token, refresh_token, expires_at, user_id)
        )
        db.commit()

    headers = {'Authorization': f'Bearer {access_token}', 'Content-Type': 'application/json'}

    message = {
        'subject': subject,
        'body': {'contentType': 'HTML', 'content': body_html},
        'toRecipients':  [{'emailAddress': {'address': to_addr}}] if to_addr else [],
        'ccRecipients':  [{'emailAddress': {'address': a}} for a in (cc_list or []) if a],
    }

    resp = _req.post('https://graph.microsoft.com/v1.0/me/messages', headers=headers, json=message)
    if resp.status_code not in (200, 201):
        return None, f'Graph API error {resp.status_code}: {resp.text[:300]}'

    draft_id = resp.json().get('id')

    if attachment_bytes and attachment_filename and draft_id:
        attach_payload = {
            '@odata.type':  '#microsoft.graph.fileAttachment',
            'name':         attachment_filename,
            'contentBytes': base64.b64encode(attachment_bytes).decode('ascii'),
            'contentType':  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        }
        _req.post(
            f'https://graph.microsoft.com/v1.0/me/messages/{draft_id}/attachments',
            headers=headers, json=attach_payload
        )

    return draft_id, None


@app.route('/pickup/<int:cid>/email/housing')
def pickup_email_housing(cid):
    """Generate the Housing History Form, save to temp file, and open in Outlook with attachment."""
    import platform
    from datetime import datetime as dt

    db = get_db()
    config = db.execute("SELECT * FROM pickup_config WHERE id=?", (cid,)).fetchone()
    if not config:
        flash('Event not found.', 'error')
        return redirect(url_for('pickup_event', cid=cid))

    pipeline = None
    if config['booking_id']:
        pipeline = db.execute(
            'SELECT * FROM ReportPipeline WHERE BookingId = ?', (config['booking_id'],)
        ).fetchone()

    sorted_dates  = sorted(json.loads(config['contracted_block'] or '{}').keys())
    org_name      = (pipeline['AccountName'] or '') if pipeline else (config['organization'] or '')
    event_name    = config['event_name'] or config['organization'] or ''
    hotel_email   = config['hotel_contact_email'] or ''
    hotel_contact = config['hotel_contact'] or ''

    if sorted_dates:
        start_str  = dt.strptime(sorted_dates[0],  '%Y-%m-%d').strftime('%m/%d/%Y')
        end_str    = dt.strptime(sorted_dates[-1], '%Y-%m-%d').strftime('%m/%d/%Y')
        date_range = f"{start_str} – {end_str}"
    else:
        date_range = ''

    first_name = hotel_contact.split(',')[-1].strip().split()[0] if hotel_contact else 'Team'
    subject    = f"Final Housing History Form for {event_name}, {org_name}, {date_range}"
    body_text  = (
        f"Dear {first_name},\n\n"
        f"Please see the attached Housing History form for {event_name}. "
        f"Please fill in all the relevant areas highlighted in yellow. "
        f"Please include the relevant pre and post days in the count. "
        f"Attach the final rooming list to your email response to me. "
        f"If you have any questions please reach out."
    )
    body_html  = body_text.replace('\n\n', '<br><br>').replace('\n', '<br>')

    user = get_current_user()

    # ── Local Mac: auto-open Outlook with form attached ────────────────────
    if platform.system() == 'Darwin':
        wb, _ = _build_housing_form_wb(config, pipeline)
        safe_name   = event_name.replace('/', '-').replace(' ', '_')[:50]
        attach_path = f'/tmp/Housing_History_{safe_name}.xlsx'
        wb.save(attach_path)
        ok, err = _open_in_outlook(cid, subject, hotel_email, [],
                                   body_html, 'housing', attachments=[attach_path])
        if ok:
            flash('Housing History email opened in Outlook with form attached — review and send.', 'success')
        else:
            flash(f'Could not open Outlook: {err}', 'error')
        return redirect(url_for('pickup_event', cid=cid))

    # ── Web: try Graph API if user has connected Microsoft account ─────────
    ms_row = db.execute(
        'SELECT user_id FROM UserMicrosoftTokens WHERE user_id = ?', (user['id'],)
    ).fetchone() if user else None

    if ms_row:
        from datetime import date as _date
        wb, _ = _build_housing_form_wb(config, pipeline)
        buf = io.BytesIO(); wb.save(buf); buf.seek(0)
        safe_name = event_name.replace('/', '-').replace(' ', '_')[:50]
        draft_id, err = _create_outlook_draft(
            user_id=user['id'], subject=subject, to_addr=hotel_email, cc_list=[],
            body_html=body_html, attachment_bytes=buf.getvalue(),
            attachment_filename=f'Housing_History_{safe_name}.xlsx',
        )
        if draft_id:
            db.execute(
                "INSERT INTO pickup_contact_log (config_id, contact_date, contact_type, notes) VALUES (?,?,?,?)",
                (cid, _date.today().isoformat(), 'email_sent',
                 f'Outlook draft created via Graph API — To: {hotel_email} (with attachment)')
            )
            db.commit()
            flash('Draft created in your Outlook Drafts folder with the form attached. Open Outlook to review and send.', 'success')
            return redirect(url_for('pickup_event', cid=cid))
        else:
            flash(f'Could not create Outlook draft: {err}', 'error')

    # ── Fallback: preview page ─────────────────────────────────────────────
    email = {'to': hotel_email, 'cc': '', 'subject': subject, 'body': body_text}
    return render_template('pickup_email_preview.html', config=config, email=email,
                           email_type='hotel', show_housing_download=True,
                           ms_connected=bool(ms_row))


@app.route('/pickup/<int:cid>/email/hotel')
def pickup_email_hotel(cid):
    """Open a hotel pickup-status email directly in Outlook."""
    import platform
    from pickup_utils import build_hotel_email

    db = get_db()
    config = db.execute("SELECT * FROM pickup_config WHERE id=?", (cid,)).fetchone()
    if not config:
        flash('Event not found.', 'error')
        return redirect(url_for('pickup_dashboard'))

    email       = build_hotel_email(config)
    hotel_email = config['hotel_contact_email'] or ''
    cc_list     = [a.strip() for a in (email.get('cc') or '').replace(';', ',').split(',') if a.strip()]
    body_html   = email.get('body', '').replace('\n', '<br>')
    user        = get_current_user()

    # ── Local Mac: auto-open Outlook ───────────────────────────────────────
    if platform.system() == 'Darwin':
        ok, err = _open_in_outlook(cid, email.get('subject', ''), hotel_email,
                                   cc_list, body_html, 'hotel')
        if ok:
            flash('Hotel email opened in Outlook — review and send when ready.', 'success')
        else:
            flash(f'Could not open Outlook: {err}', 'error')
        return redirect(url_for('pickup_event', cid=cid))

    # ── Web: try Graph API if user has connected Microsoft account ─────────
    ms_row = db.execute(
        'SELECT user_id FROM UserMicrosoftTokens WHERE user_id = ?', (user['id'],)
    ).fetchone() if user else None

    if ms_row:
        from datetime import date as _date
        draft_id, err = _create_outlook_draft(
            user_id=user['id'], subject=email.get('subject', ''),
            to_addr=hotel_email, cc_list=cc_list, body_html=body_html,
        )
        if draft_id:
            db.execute(
                "INSERT INTO pickup_contact_log (config_id, contact_date, contact_type, notes) VALUES (?,?,?,?)",
                (cid, _date.today().isoformat(), 'email_sent',
                 f'Outlook draft created via Graph API — To: {hotel_email}')
            )
            db.commit()
            flash('Draft created in your Outlook Drafts folder. Open Outlook to review and send.', 'success')
            return redirect(url_for('pickup_event', cid=cid))
        else:
            flash(f'Could not create Outlook draft: {err}', 'error')

    # ── Fallback: preview page ─────────────────────────────────────────────
    return render_template('pickup_email_preview.html', config=config, email=email,
                           email_type='hotel', ms_connected=bool(ms_row))


@app.route('/pickup/<int:cid>/email/hotel/rate-issue')
def pickup_email_hotel_rate_issue(cid):
    """Open a hotel rate-parity issue email in Outlook."""
    import platform
    from pickup_utils import build_hotel_rate_issue_email

    db = get_db()
    config = db.execute("SELECT * FROM pickup_config WHERE id=?", (cid,)).fetchone()
    if not config:
        flash('Event not found.', 'error')
        return redirect(url_for('pickup_dashboard'))

    # Get the most recent real weekly entry for OTA data
    last_real = db.execute(
        """SELECT * FROM pickup_weekly WHERE config_id=?
           AND (total_rooms IS NOT NULL AND total_rooms > 0)
           AND (label IS NULL OR (label NOT LIKE 'ASAS %' AND label NOT LIKE 'Historical%'
                                  AND label NOT LIKE '%Final%'))
           ORDER BY report_date DESC LIMIT 1""",
        (cid,)
    ).fetchone()

    if not last_real or not last_real['ota_rate']:
        flash('No OTA rate found in the latest weekly entry.', 'error')
        return redirect(url_for('pickup_event', cid=cid))

    # Use weekly entry's ota_url first, fall back to config-level ota_url
    ota_url = (last_real['ota_url'] or config['ota_url'] or '').strip() or None
    email = build_hotel_rate_issue_email(config, last_real['ota_rate'], ota_url)
    hotel_email = config['hotel_contact_email'] or ''
    cc_list     = [a.strip() for a in (email.get('cc') or '').replace(';', ',').split(',') if a.strip()]
    body_html   = email.get('body_html') or email.get('body', '').replace('\n', '<br>')
    user        = get_current_user()

    if platform.system() == 'Darwin':
        ok, err = _open_in_outlook(cid, email.get('subject', ''), hotel_email,
                                   cc_list, body_html, 'hotel_rate_issue')
        if ok:
            flash('Rate issue email opened in Outlook — review and send when ready.', 'success')
        else:
            flash(f'Could not open Outlook: {err}', 'error')
        return redirect(url_for('pickup_event', cid=cid))

    ms_row = db.execute(
        'SELECT user_id FROM UserMicrosoftTokens WHERE user_id = ?', (user['id'],)
    ).fetchone() if user else None

    if ms_row:
        from datetime import date as _date
        draft_id, err = _create_outlook_draft(
            user_id=user['id'], subject=email.get('subject', ''),
            to_addr=hotel_email, cc_list=cc_list, body_html=body_html,
        )
        if draft_id:
            flash('Rate issue draft created in your Outlook Drafts folder.', 'success')
            return redirect(url_for('pickup_event', cid=cid))
        else:
            flash(f'Could not create Outlook draft: {err}', 'error')

    return render_template('pickup_email_preview.html', config=config, email=email,
                           email_type='hotel_rate_issue', ms_connected=bool(ms_row))


@app.route('/pickup/<int:cid>/email/client')
def pickup_email_client(cid):
    import platform
    from datetime import date as _date
    from pickup_utils import build_client_email, _build_cc_recipients
    db = get_db()
    config = db.execute("SELECT * FROM pickup_config WHERE id=?", (cid,)).fetchone()
    last = db.execute("SELECT * FROM pickup_weekly WHERE config_id=? ORDER BY report_date DESC LIMIT 1", (cid,)).fetchone()
    weekly_list = db.execute("SELECT * FROM pickup_weekly WHERE config_id=? ORDER BY report_date DESC", (cid,)).fetchall()
    rl = db.execute("SELECT id, reconciliation_status, filename, total_guests, upload_date FROM pickup_rooming_list WHERE config_id=? ORDER BY upload_date DESC LIMIT 1", (cid,)).fetchone()
    if not config:
        flash('Event not found.', 'error')
        return redirect(url_for('pickup_dashboard'))
    rl_status    = rl['reconciliation_status'] if rl else None
    weekly_dicts = [dict(w) for w in weekly_list]
    email        = build_client_email(config, last, rl_status, weekly_list=weekly_dicts)

    user   = get_current_user()
    ms_row = db.execute('SELECT user_id FROM UserMicrosoftTokens WHERE user_id=?',
                        (user['id'],)).fetchone() if user else None

    # ── Web: try Graph API if user has connected Microsoft account ─────────
    if platform.system() != 'Darwin' and ms_row:
        html_body     = email.get('html_body') or email.get('body', '').replace('\n', '<br>')
        to_addr       = (email.get('to') or '').strip()
        cc_recipients = _build_cc_recipients(config)
        cc_list       = [r['email'] for r in cc_recipients if r.get('email')]
        draft_id, err = _create_outlook_draft(
            user_id=user['id'], subject=email.get('subject', ''),
            to_addr=to_addr, cc_list=cc_list, body_html=html_body,
        )
        if draft_id:
            db.execute(
                "INSERT INTO pickup_contact_log (config_id, contact_date, contact_type, notes) VALUES (?,?,?,?)",
                (cid, _date.today().isoformat(), 'email_sent',
                 f'Client draft created via Graph API — To: {to_addr}')
            )
            db.commit()
            flash('Draft created in your Outlook Drafts folder. Open Outlook to review and send.', 'success')
            return redirect(url_for('pickup_event', cid=cid))
        else:
            flash(f'Could not create Outlook draft: {err}', 'error')

    return render_template('pickup_email_preview.html', config=config, email=email,
                           email_type='client', rooming_list=rl,
                           ms_connected=bool(ms_row))


@app.route('/pickup/<int:cid>/email/client/launch-outlook')
def pickup_email_client_launch_outlook(cid):
    import subprocess, tempfile, platform
    from datetime import date as _date
    from pickup_utils import build_client_email, _build_cc_recipients
    db = get_db()
    config = db.execute("SELECT * FROM pickup_config WHERE id=?", (cid,)).fetchone()
    last = db.execute("SELECT * FROM pickup_weekly WHERE config_id=? ORDER BY report_date DESC LIMIT 1", (cid,)).fetchone()
    weekly_list = db.execute("SELECT * FROM pickup_weekly WHERE config_id=? ORDER BY report_date DESC", (cid,)).fetchall()
    rl = db.execute("SELECT reconciliation_status FROM pickup_rooming_list WHERE config_id=? ORDER BY upload_date DESC LIMIT 1", (cid,)).fetchone()
    if not config:
        flash('Event not found.', 'error')
        return redirect(url_for('pickup_dashboard'))
    rl_status    = rl['reconciliation_status'] if rl else None
    weekly_dicts = [dict(w) for w in weekly_list]
    email         = build_client_email(config, last, rl_status, weekly_list=weekly_dicts)
    html_body     = email.get('html_body', '')
    subject       = (email.get('subject') or '').strip()
    to_addr       = (email.get('to') or '').strip()
    cc_recipients = _build_cc_recipients(config)
    cc_list       = [r['email'] for r in cc_recipients if r.get('email')]

    # ── Web: Graph API draft ───────────────────────────────────────────────
    if platform.system() != 'Darwin':
        user   = get_current_user()
        ms_row = db.execute('SELECT user_id FROM UserMicrosoftTokens WHERE user_id=?',
                            (user['id'],)).fetchone() if user else None
        if ms_row:
            draft_id, err = _create_outlook_draft(
                user_id=user['id'], subject=subject,
                to_addr=to_addr, cc_list=cc_list, body_html=html_body,
            )
            if draft_id:
                db.execute(
                    "INSERT INTO pickup_contact_log (config_id, contact_date, contact_type, notes) VALUES (?,?,?,?)",
                    (cid, _date.today().isoformat(), 'email_sent',
                     f'Client draft created via Graph API — To: {to_addr}')
                )
                db.commit()
                flash('Draft created in your Outlook Drafts folder. Open Outlook to review and send.', 'success')
                return redirect(url_for('pickup_event', cid=cid))
            else:
                flash(f'Could not create Outlook draft: {err}', 'error')
        return redirect(url_for('pickup_email_client', cid=cid))

    # ── Local Mac: JXA clipboard + AppleScript ────────────────────────────
    def write_tmp(content, suffix):
        t = tempfile.NamedTemporaryFile(mode='w', suffix=suffix, delete=False, encoding='utf-8')
        t.write(content); t.close(); return t.name

    def esc(s):
        return s.replace('\\', '\\\\').replace('"', '\\"')

    tmp = write_tmp(html_body, '.html')
    clip_jxa = f"""ObjC.import('AppKit'); ObjC.import('Foundation');
var nsStr = $.NSString.alloc.initWithContentsOfFileEncodingError('{tmp}', $.NSUTF8StringEncoding, null);
var html = ObjC.unwrap(nsStr);
var pb = $.NSPasteboard.generalPasteboard; pb.clearContents;
pb.setStringForType($.NSString.alloc.initWithUTF8String(html), $.NSPasteboardTypeHTML);"""

    to_line  = (f'make new to recipient at theMsg with properties {{email address:{{name:"", address:"{esc(to_addr)}"}}}}' if to_addr else '')
    cc_lines = '\n    '.join(f'make new cc recipient at theMsg with properties {{email address:{{name:"{esc(r["name"])}", address:"{esc(r["email"])}"}}}}' for r in cc_recipients)
    outlook_script = (
        'tell application "Microsoft Outlook"\n'
        f'    set theMsg to make new outgoing message with properties {{subject:"{esc(subject)}"}}\n'
        + (f'    {to_line}\n' if to_line else '')
        + (f'    {cc_lines}\n' if cc_lines else '')
        + '    open theMsg\n    activate\nend tell\n'
    )
    try:
        clip_path    = write_tmp(clip_jxa, '.js')
        outlook_path = write_tmp(outlook_script, '.applescript')
        subprocess.run(['/usr/bin/osascript', '-l', 'JavaScript', clip_path], check=True)
        subprocess.Popen(['/usr/bin/osascript', outlook_path])
    except Exception as exc:
        flash(f'Could not launch Outlook: {exc}', 'error')
        return redirect(url_for('pickup_email_client', cid=cid))
    return redirect(url_for('pickup_email_client_paste', cid=cid))


@app.route('/pickup/<int:cid>/email/client/paste-instructions')
def pickup_email_client_paste(cid):
    db = get_db()
    config = db.execute("SELECT * FROM pickup_config WHERE id=?", (cid,)).fetchone()
    return render_template('pickup_outlook_paste.html', config=config)


@app.route('/pickup/<int:cid>/export-row')
def pickup_export_row(cid):
    from pickup_utils import format_export_row
    db = get_db()
    config = db.execute("SELECT * FROM pickup_config WHERE id=?", (cid,)).fetchone()
    last   = db.execute("SELECT * FROM pickup_weekly WHERE config_id=? ORDER BY report_date DESC LIMIT 1", (cid,)).fetchone()
    if not config or not last:
        flash('No data to export.', 'error')
        return redirect(url_for('pickup_event', cid=cid))
    row_text = format_export_row(config, last)
    return render_template('pickup_export_row.html', config=config, row_text=row_text)


@app.route('/pickup/<int:cid>/hhr/download')
def pickup_hhr_download(cid):
    import io as _io
    db = get_db()
    config = db.execute('SELECT * FROM pickup_config WHERE id=?', (cid,)).fetchone()
    if not config:
        flash('Event not found.', 'error')
        return redirect(url_for('pickup_event', cid=cid))
    hhr = db.execute(
        'SELECT filename, file_data FROM housing_history_files WHERE booking_id=? ORDER BY id DESC LIMIT 1',
        (config['booking_id'],)
    ).fetchone()
    if not hhr or not hhr['file_data']:
        flash('No Housing History Report on file.', 'error')
        return redirect(url_for('pickup_event', cid=cid))
    file_bytes = bytes(hhr['file_data'])
    try:
        from pickup_utils import strip_hhr_commission_rows, clean_hhr_for_client
        file_bytes = clean_hhr_for_client(strip_hhr_commission_rows(file_bytes))
    except Exception:
        pass
    return send_file(_io.BytesIO(file_bytes),
                     download_name=hhr['filename'] or 'housing_history.xlsx',
                     as_attachment=True)


def _get_post_report_data(cid):
    """Assemble stats and config_dict for the Post Report from correct DB sources."""
    db = get_db()
    config = db.execute('SELECT * FROM pickup_config WHERE id=?', (cid,)).fetchone()
    if not config:
        return None, None, None

    fh = db.execute(
        "SELECT * FROM pickup_weekly WHERE config_id=? AND label='Final History' ORDER BY id DESC LIMIT 1",
        (cid,)
    ).fetchone()
    if not fh:
        return config, None, None

    hhr_row = db.execute(
        'SELECT filename, file_data FROM housing_history_files WHERE booking_id=? ORDER BY id DESC LIMIT 1',
        (config['booking_id'],)
    ).fetchone()

    pickup_row = db.execute(
        'SELECT ActualPickup, TotalRevenue FROM Pickup WHERE BookingID=?',
        (config['booking_id'],)
    ).fetchone()

    hhr_stats = {}
    if hhr_row and hhr_row['file_data']:
        try:
            from pickup_utils import parse_hhr_excel
            hhr_stats = parse_hhr_excel(bytes(hhr_row['file_data']))
        except Exception:
            pass

    block = json.loads(config['contracted_block'] or '{}')
    contracted_total = sum(block.values()) or hhr_stats.get('contracted_total') or 0
    pbn = json.loads(fh['pickup_by_night'] or '{}') if fh['pickup_by_night'] else {}
    final_total_pickup = (fh['total_rooms'] or hhr_stats.get('final_total_pickup')
                          or (float(pickup_row['ActualPickup']) if pickup_row and pickup_row['ActualPickup'] else 0))

    stats = {
        'organization':         config['organization'] or hhr_stats.get('organization', ''),
        'hotel':                config['hotel']        or hhr_stats.get('hotel', ''),
        'event_name':           config['event_name']   or hhr_stats.get('event_name', ''),
        'contracted_total':     contracted_total,
        'contracted_block':     block,
        'contracted_rate':      hhr_stats.get('contracted_rate') or config['contracted_rate'],
        'final_total_pickup':   final_total_pickup,
        'final_pickup_by_night': pbn,
        'pct_of_block':         fh['pct_of_block'] or hhr_stats.get('pct_of_block'),
        'pct_of_attrition':     fh['pct_of_attrition'],
        'room_revenue':         hhr_stats.get('room_revenue') or (float(pickup_row['TotalRevenue']) if pickup_row and pickup_row['TotalRevenue'] else None),
        'fb_revenue':           hhr_stats.get('fb_revenue'),
        'audit_pickup':         hhr_stats.get('audit_pickup'),
        'no_shows':             hhr_stats.get('no_shows'),
        'cancellations':        hhr_stats.get('cancellations'),
        'hotel_approver':       hhr_stats.get('hotel_approver'),
        'hotel_approver_email': hhr_stats.get('hotel_approver_email'),
    }

    config_dict = dict(config)
    # Name the attachment after the event
    import re as _re
    _event = (config['event_name'] or config['organization'] or 'Housing History Report').strip()
    _safe  = _re.sub(r'[\\/*?:"<>|]', '', _event)   # strip chars illegal in filenames
    config_dict['hhr_filename'] = f"{_safe} — Housing History Report.xlsx"

    # Strip commission rows before client delivery
    raw_bytes = bytes(hhr_row['file_data']) if (hhr_row and hhr_row['file_data']) else None
    if raw_bytes:
        try:
            from pickup_utils import strip_hhr_commission_rows, clean_hhr_for_client
            raw_bytes = clean_hhr_for_client(strip_hhr_commission_rows(raw_bytes))
        except Exception:
            pass  # fall back to unstripped file if anything goes wrong
    config_dict['_hhr_file_data'] = raw_bytes

    return config_dict, stats, fh


@app.route('/pickup/<int:cid>/email/post-report')
def pickup_email_post_report(cid):
    config_dict, stats, fh = _get_post_report_data(cid)
    if config_dict is None:
        flash('Event not found.', 'error')
        return redirect(url_for('pickup_dashboard'))
    if stats is None:
        flash('No Final History on file — import the HHR first via Import → HHR.', 'warning')
        return redirect(url_for('pickup_event', cid=cid))
    email = _build_post_report_email(config_dict, stats)
    return render_template('pickup_post_report_email.html', config=config_dict, stats=stats, email=email)


@app.route('/pickup/<int:cid>/email/post-report/launch-outlook')
def pickup_email_post_report_outlook(cid):
    import subprocess, tempfile, platform
    from datetime import date as _date

    config_dict, stats, fh = _get_post_report_data(cid)
    if config_dict is None:
        flash('Event not found.', 'error')
        return redirect(url_for('pickup_dashboard'))
    if stats is None:
        flash('No Final History on file — import the HHR first via Import → HHR.', 'warning')
        return redirect(url_for('pickup_event', cid=cid))

    # ── Web: Graph API draft with HHR attachment ───────────────────────────
    if platform.system() not in ('Darwin', 'Windows'):
        user   = get_current_user()
        db     = get_db()
        ms_row = db.execute('SELECT user_id FROM UserMicrosoftTokens WHERE user_id=?',
                            (user['id'],)).fetchone() if user else None
        if ms_row:
            email        = _build_post_report_email(config_dict, stats)
            html_body    = email.get('html_body', '')
            subject      = email.get('subject', '')
            to_addr      = email.get('to', '')
            file_data    = config_dict.get('_hhr_file_data')
            hhr_filename = config_dict.get('hhr_filename') or 'Housing History Report.xlsx'
            from pickup_utils import _build_cc_recipients
            cc_list = [r['email'] for r in _build_cc_recipients(config_dict) if r.get('email')]
            draft_id, err = _create_outlook_draft(
                user_id=user['id'], subject=subject,
                to_addr=to_addr, cc_list=cc_list, body_html=html_body,
                attachment_bytes=file_data, attachment_filename=hhr_filename,
            )
            if draft_id:
                flash('Post Report draft created in your Outlook Drafts folder with HHR attached. Open Outlook to review and send.', 'success')
                return redirect(url_for('pickup_event', cid=config_dict['id']))
            else:
                flash(f'Could not create Outlook draft: {err}', 'error')
        else:
            flash('Connect your Microsoft account (My Account) to create Outlook drafts from the website.', 'warning')
        return redirect(url_for('pickup_email_post_report', cid=config_dict['id']))

    email     = _build_post_report_email(config_dict, stats)
    html_body = email.get('html_body', '')
    subject   = email.get('subject', '')
    to_addr   = email.get('to', '')

    from pickup_utils import _build_cc_recipients
    cc_recipients = _build_cc_recipients(config_dict)

    file_data    = config_dict.get('_hhr_file_data')
    hhr_filename = config_dict.get('hhr_filename') or 'Housing History Report.xlsx'

    def write_tmp(content, suffix, mode='w', encoding='utf-8'):
        t = tempfile.NamedTemporaryFile(mode=mode, suffix=suffix, delete=False,
                                        encoding=encoding if mode == 'w' else None)
        t.write(content); t.close(); return t.name

    def write_named_att(data, filename):
        """Write attachment bytes to a temp dir using the proper display filename."""
        import os as _os
        d = tempfile.mkdtemp()
        p = _os.path.join(d, filename)
        with open(p, 'wb') as f:
            f.write(data)
        return p

    try:
        if platform.system() == 'Windows':
            # ── Windows: PowerShell + Outlook COM ─────────────────────────────
            tmp_html = write_tmp(html_body, '.html')

            att_path = ''
            if file_data:
                att_path = write_named_att(file_data, hhr_filename).replace('\\', '\\\\')

            cc_str = '; '.join(r['email'] for r in cc_recipients if r.get('email'))

            ps_lines = [
                '$html = Get-Content -Path \'' + tmp_html.replace("'", "''") + '\' -Raw -Encoding UTF8',
                '$ol = New-Object -ComObject Outlook.Application',
                '$mail = $ol.CreateItem(0)',
                f'$mail.Subject = \'{subject.replace(chr(39), chr(39)*2)}\'',
                '$mail.HTMLBody = $html',
            ]
            if to_addr:
                ps_lines.append(f'$mail.To = \'{to_addr.replace(chr(39), chr(39)*2)}\'')
            if cc_str:
                ps_lines.append(f'$mail.CC = \'{cc_str.replace(chr(39), chr(39)*2)}\'')
            if att_path:
                ps_lines.append(f'$mail.Attachments.Add(\'{att_path}\') | Out-Null')
            ps_lines.append('$mail.Display()')

            ps_script = '\r\n'.join(ps_lines)
            ps_path   = write_tmp(ps_script, '.ps1')
            subprocess.Popen(['powershell.exe', '-ExecutionPolicy', 'Bypass',
                               '-WindowStyle', 'Hidden', '-File', ps_path])

        else:
            # ── macOS: AppleScript + System Events auto-paste ─────────────────
            def esc(s):
                return s.replace('\\', '\\\\').replace('"', '\\"')

            tmp_html = write_tmp(html_body, '.html')
            clip_jxa = (
                "ObjC.import('AppKit'); ObjC.import('Foundation');\n"
                f"var nsStr = $.NSString.alloc.initWithContentsOfFileEncodingError('{tmp_html}', $.NSUTF8StringEncoding, null);\n"
                "var html = ObjC.unwrap(nsStr);\n"
                "var pb = $.NSPasteboard.generalPasteboard; pb.clearContents;\n"
                "pb.setStringForType($.NSString.alloc.initWithUTF8String(html), $.NSPasteboardTypeHTML);"
            )

            attach_line = ''
            if file_data:
                att_path = write_named_att(file_data, hhr_filename)
                attach_line = f'make new attachment at theMsg with properties {{file:POSIX file "{att_path}"}}'

            to_line = (
                f'make new to recipient at theMsg with properties {{email address:{{name:"", address:"{esc(to_addr)}"}}}}'
                if to_addr else ''
            )
            cc_lines = '\n    '.join(
                f'make new cc recipient at theMsg with properties {{email address:{{name:"{esc(r["name"])}", address:"{esc(r["email"])}"}}}}'
                for r in cc_recipients
            )

            outlook_script = (
                'tell application "Microsoft Outlook"\n'
                f'    set theMsg to make new outgoing message with properties {{subject:"{esc(subject)}"}}\n'
                + (f'    {to_line}\n' if to_line else '')
                + (f'    {cc_lines}\n' if cc_lines else '')
                + (f'    {attach_line}\n' if attach_line else '')
                + '    open theMsg\n'
                + '    activate\n'
                + 'end tell\n'
                + 'delay 2\n'
                + 'tell application "System Events"\n'
                + '    keystroke "v" using {command down}\n'
                + 'end tell\n'
            )

            clip_path    = write_tmp(clip_jxa, '.js')
            outlook_path = write_tmp(outlook_script, '.applescript')
            subprocess.run(['osascript', '-l', 'JavaScript', clip_path], check=True)
            subprocess.Popen(['osascript', outlook_path])

    except Exception as exc:
        flash(f'Could not launch Outlook: {exc}', 'error')
        return redirect(url_for('pickup_event', cid=cid))

    flash('Post Report email opened in Outlook.', 'success')
    return redirect(url_for('pickup_event', cid=cid))


def _build_post_report_email(config, stats):
    """Build the post-event housing history email dict."""
    from datetime import date as _date
    org        = config['organization'] or stats.get('organization', '')
    event_name = config['event_name']   or stats.get('event_name', '')
    hotel      = config['hotel']        or stats.get('hotel', '')
    to_addr    = config['group_contact_email'] or ''

    ct       = stats.get('contracted_total') or 0
    fp       = stats.get('final_total_pickup') or 0
    pct      = stats.get('pct_of_block')
    rate     = stats.get('contracted_rate')
    rev      = stats.get('room_revenue')
    fb_rev   = stats.get('fb_revenue')
    audit    = stats.get('audit_pickup') or 0
    no_shows = stats.get('no_shows') or 0
    cancels  = stats.get('cancellations') or 0

    attrition_rooms = round(ct * float(config['attrition_pct'] or 0)) if ct and config.get('attrition_pct') else None
    pct_attrition   = round(fp / attrition_rooms * 100, 1) if attrition_rooms and fp else None
    if not pct_attrition:
        pct_attrition_val = stats.get('pct_of_attrition')
        if pct_attrition_val:
            pct_attrition = pct_attrition_val

    subject = f'{org} — {event_name} | Final Housing History Report'

    _dow = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']

    block_nights  = stats.get('contracted_block', {})
    pickup_nights = stats.get('final_pickup_by_night', {})
    all_dates = sorted(set(list(block_nights.keys()) + list(pickup_nights.keys())))
    night_rows = ''
    for d in all_dates:
        b = block_nights.get(d, 0)
        p = pickup_nights.get(d, 0)
        diff = p - b
        diff_str   = (f'+{diff}' if diff > 0 else str(diff)) if diff != 0 else '—'
        diff_color = '#16a34a' if diff > 0 else ('#dc2626' if diff < 0 else '#6b7280')
        try:
            dow = _dow[_date.fromisoformat(d).weekday()]
        except Exception:
            dow = ''
        date_label = f'{d[5:].replace("-", "/")} ({dow})' if dow else d[5:].replace('-', '/')
        night_rows += (
            f'<tr>'
            f'<td style="padding:4px 10px;border-bottom:1px solid #e5e7eb">{date_label}</td>'
            f'<td style="padding:4px 10px;text-align:center;border-bottom:1px solid #e5e7eb">{b or "—"}</td>'
            f'<td style="padding:4px 10px;text-align:center;border-bottom:1px solid #e5e7eb">{p or "—"}</td>'
            f'<td style="padding:4px 10px;text-align:center;border-bottom:1px solid #e5e7eb;color:{diff_color}">{diff_str}</td>'
            f'</tr>'
        )

    pct_str     = f'{pct:.1f}%' if pct is not None else 'N/A'
    pct_att_str = f'{pct_attrition:.1f}%' if pct_attrition is not None else 'N/A'
    rate_str    = f'${rate:,.2f}' if rate else 'N/A'
    rev_str     = f'${rev:,.2f}' if rev else 'N/A'
    fb_str      = f'${fb_rev:,.2f}' if fb_rev else 'N/A'
    pct_color   = '#16a34a' if (pct or 0) >= 100 else '#d97706'
    att_color   = '#16a34a' if (pct_attrition or 0) >= 100 else '#d97706'

    html_body = f'''<div style="font-family:Arial,sans-serif;max-width:640px;color:#1f2937">
<p>Hi {config.get("group_contact") or ""},</p>
<p>Please find attached the final housing history report for <strong>{event_name}</strong> at <strong>{hotel}</strong>. Below is a summary of the pickup performance.</p>

<table style="width:100%;border-collapse:collapse;margin:16px 0;background:#f9fafb;border-radius:8px;overflow:hidden">
  <tr style="background:#1a3a5c;color:#fff">
    <th style="padding:8px 10px;text-align:left" colspan="2">Pickup Performance Summary</th>
  </tr>
  <tr><td style="padding:6px 10px;border-bottom:1px solid #e5e7eb">Contracted Block</td>
      <td style="padding:6px 10px;border-bottom:1px solid #e5e7eb;font-weight:600">{int(ct):,} room nights</td></tr>
  <tr><td style="padding:6px 10px;border-bottom:1px solid #e5e7eb">Final Total Pickup</td>
      <td style="padding:6px 10px;border-bottom:1px solid #e5e7eb;font-weight:600">{int(fp):,} room nights</td></tr>
  <tr><td style="padding:6px 10px;border-bottom:1px solid #e5e7eb">% of Contracted Block</td>
      <td style="padding:6px 10px;border-bottom:1px solid #e5e7eb;font-weight:600;color:{pct_color}">{pct_str}</td></tr>
  {'<tr><td style="padding:6px 10px;border-bottom:1px solid #e5e7eb">% of Attrition Commitment</td><td style="padding:6px 10px;border-bottom:1px solid #e5e7eb;font-weight:600;color:' + att_color + '">' + pct_att_str + '</td></tr>' if pct_attrition else ''}
  <tr><td style="padding:6px 10px;border-bottom:1px solid #e5e7eb">Contracted Rate</td>
      <td style="padding:6px 10px;border-bottom:1px solid #e5e7eb">{rate_str}</td></tr>
  <tr><td style="padding:6px 10px;border-bottom:1px solid #e5e7eb">Total Room Revenue</td>
      <td style="padding:6px 10px;border-bottom:1px solid #e5e7eb">{rev_str}</td></tr>
  {'<tr><td style="padding:6px 10px;border-bottom:1px solid #e5e7eb">Food &amp; Beverage Revenue</td><td style="padding:6px 10px;border-bottom:1px solid #e5e7eb">' + fb_str + '</td></tr>' if fb_rev else ''}
  {'<tr><td style="padding:6px 10px;border-bottom:1px solid #e5e7eb">Audit Pickup (Outside Block)</td><td style="padding:6px 10px;border-bottom:1px solid #e5e7eb">' + f'{int(audit):,} rooms' + '</td></tr>' if audit else ''}
  {'<tr><td style="padding:6px 10px">No Shows / Cancellations</td><td style="padding:6px 10px">' + f'{int(no_shows):,} / {int(cancels):,}' + '</td></tr>' if (no_shows or cancels) else ''}
</table>
{'<table style="width:100%;border-collapse:collapse;margin:16px 0"><thead><tr style="background:#e5e7eb"><th style="padding:6px 10px;text-align:left">Night</th><th style="padding:6px 10px;text-align:center">Block</th><th style="padding:6px 10px;text-align:center">Pickup</th><th style="padding:6px 10px;text-align:center">+/−</th></tr></thead><tbody>' + night_rows + '</tbody></table>' if night_rows else ''}
<p>The full Housing History Report is attached. Please don't hesitate to reach out with any questions.</p>
<p>Best regards,</p>
</div>'''

    return {'to': to_addr, 'subject': subject, 'html_body': html_body}


@app.route('/pickup/<int:cid>/archive', methods=['POST'])
def pickup_archive(cid):
    db = get_db()
    db.execute("UPDATE pickup_config SET status='archived' WHERE id=?", (cid,))
    db.commit()
    flash('Event archived.', 'success')
    return redirect(url_for('pickup_dashboard'))


@app.route('/pickup/<int:cid>/unarchive', methods=['POST'])
def pickup_unarchive(cid):
    db = get_db()
    db.execute("UPDATE pickup_config SET status='active' WHERE id=?", (cid,))
    db.commit()
    flash('Event restored to active.', 'success')
    return redirect(url_for('pickup_dashboard'))


@app.route('/pickup/<int:cid>/toggle-current', methods=['POST'])
def pickup_toggle_current(cid):
    db = get_db()
    row = db.execute("SELECT force_current FROM pickup_config WHERE id=?", (cid,)).fetchone()
    new_val = 0 if (row and row[0]) else 1
    db.execute("UPDATE pickup_config SET force_current=? WHERE id=?", (new_val, cid))
    db.commit()
    return redirect(url_for('pickup_dashboard'))


@app.route('/pickup/<int:cid>/promote-past', methods=['POST'])
def pickup_promote_past(cid):
    db = get_db()
    row = db.execute("SELECT force_past FROM pickup_config WHERE id=?", (cid,)).fetchone()
    new_val = 0 if (row and row['force_past']) else 1
    db.execute("UPDATE pickup_config SET force_past=? WHERE id=?", (new_val, cid))
    db.commit()
    return redirect(url_for('pickup_dashboard'))


@app.route('/pickup/<int:cid>/contact-log/add', methods=['POST'])
def pickup_contact_log_add(cid):
    contact_date = request.form.get('contact_date', '').strip()
    contact_type = request.form.get('contact_type', 'email_sent')
    notes        = request.form.get('notes', '').strip()
    if contact_date:
        db = get_db()
        db.execute("INSERT INTO pickup_contact_log (config_id, contact_date, contact_type, notes) VALUES (?,?,?,?)",
                   (cid, contact_date, contact_type, notes or None))
        db.commit()
    return redirect(url_for('pickup_event', cid=cid))


@app.route('/pickup/<int:cid>/contact-log/<int:log_id>/delete', methods=['POST'])
def pickup_contact_log_delete(cid, log_id):
    db = get_db()
    db.execute("DELETE FROM pickup_contact_log WHERE id=? AND config_id=?", (log_id, cid))
    db.commit()
    return redirect(url_for('pickup_event', cid=cid))


@app.route('/pickup/<int:cid>/log-email-sent', methods=['POST'])
def pickup_log_email_sent(cid):
    from datetime import date as _date
    db = get_db()
    db.execute("INSERT INTO pickup_contact_log (config_id, contact_date, contact_type, notes) VALUES (?,?,?,?)",
               (cid, _date.today().isoformat(), 'email_sent', 'Logged automatically on Open in Mail App'))
    db.commit()
    return jsonify({'ok': True}), 200, {'Content-Type': 'application/json'}


@app.route('/pickup/<int:cid>/quick-respond', methods=['POST'])
def pickup_quick_respond(cid):
    from datetime import date as _date
    db = get_db()
    existing = db.execute("SELECT id FROM pickup_contact_log WHERE config_id=? AND contact_type='responded' LIMIT 1", (cid,)).fetchone()
    today = _date.today().isoformat()
    if not existing:
        db.execute("INSERT INTO pickup_contact_log (config_id, contact_date, contact_type, notes) VALUES (?,?,?,?)",
                   (cid, today, 'responded', 'Marked responded from dashboard'))
        db.commit()
    if request.headers.get('X-Requested-With') == 'fetch':
        return jsonify({'ok': True, 'date': today})
    return redirect(url_for('pickup_dashboard'))


@app.route('/pickup/<int:cid>/undo-respond', methods=['POST'])
def pickup_undo_respond(cid):
    db = get_db()
    db.execute("DELETE FROM pickup_contact_log WHERE config_id=? AND contact_type='responded'", (cid,))
    db.commit()
    if request.headers.get('X-Requested-With') == 'fetch':
        return jsonify({'ok': True})
    return redirect(url_for('pickup_dashboard'))


@app.route('/pickup/<int:cid>/rooming-list/debug-file', methods=['POST'])
def pickup_debug_file(cid):
    f = request.files.get('rooming_pdf')
    if not f:
        return 'No file uploaded', 400
    ext = f.filename.rsplit('.', 1)[-1].lower() if '.' in f.filename else ''
    raw_bytes = f.read()
    try:
        if ext == 'pdf':
            import pdfplumber, io as _io
            pages = []
            with pdfplumber.open(_io.BytesIO(raw_bytes)) as pdf:
                for i, page in enumerate(pdf.pages):
                    pages.append(f'=== PAGE {i+1} ===\n' + (page.extract_text() or '(no text)'))
            raw = '\n'.join(pages)
        elif ext == 'csv':
            raw = raw_bytes.decode('utf-8-sig', errors='replace')
        else:
            raw = f'Binary file ({len(raw_bytes)} bytes) — cannot display'
        return f'<pre style="font-size:.8rem;white-space:pre-wrap;word-break:break-all">{raw[:10000]}</pre>'
    except Exception as e:
        return f'Error: {e}', 500


# ── Contract Templates ────────────────────────────────────────────────────────

@app.route('/contracts/templates')
def contract_templates():
    user = get_current_user()
    if not has_permission(user, 'contracts'):
        flash('Access denied.', 'error')
        return redirect(url_for('index'))
    db = get_db()
    templates = db.execute(
        "SELECT * FROM contract_template WHERE is_active=1 ORDER BY chain, template_name"
    ).fetchall()
    from collections import defaultdict
    by_chain = defaultdict(list)
    for t in templates:
        by_chain[t['chain']].append(t)
    return render_template('contract_templates.html', by_chain=dict(by_chain))


@app.route('/contracts/templates/upload', methods=['GET', 'POST'])
def contract_template_upload():
    user = get_current_user()
    if not has_permission(user, 'contracts'):
        flash('Access denied.', 'error')
        return redirect(url_for('index'))
    from pickup_utils import extract_template_metadata
    if request.method == 'GET':
        return render_template('contract_template_upload.html')
    f = request.files.get('template_file')
    if not f or not f.filename:
        flash('No file selected.', 'error')
        return redirect(url_for('contract_templates'))
    file_bytes = f.read()
    chain = request.form.get('chain', '').strip()
    template_name = request.form.get('template_name', '').strip() or f.filename
    description = request.form.get('description', '').strip() or None
    try:
        meta = extract_template_metadata(file_bytes)
    except Exception as e:
        meta = {'merge_fields': [], 'sections': []}
        flash(f'Warning: could not parse template metadata ({e}).', 'warning')
    db = get_db()
    db.execute('''
        INSERT INTO contract_template
            (chain, template_name, description, filename, file_data, sections, merge_fields)
        VALUES (?,?,?,?,?,?,?)
    ''', (chain, template_name, description, f.filename, file_bytes,
          json.dumps(meta['sections']), json.dumps(meta['merge_fields'])))
    db.commit()
    flash(f'Template "{template_name}" uploaded ({len(meta["sections"])} sections, '
          f'{len(meta["merge_fields"])} merge fields).', 'success')
    return redirect(url_for('contract_templates'))


@app.route('/contracts/templates/<int:tid>/download')
def contract_template_download(tid):
    user = get_current_user()
    if not has_permission(user, 'contracts'):
        flash('Access denied.', 'error')
        return redirect(url_for('index'))
    import io as _io
    db = get_db()
    tmpl = db.execute("SELECT * FROM contract_template WHERE id=?", (tid,)).fetchone()
    if not tmpl or not tmpl['file_data']:
        flash('Template not found.', 'error')
        return redirect(url_for('contract_templates'))
    return send_file(_io.BytesIO(tmpl['file_data']),
                     download_name=tmpl['filename'] or 'template.docx',
                     as_attachment=True)


@app.route('/contracts/templates/<int:tid>/delete', methods=['POST'])
def contract_template_delete(tid):
    user = get_current_user()
    if not has_permission(user, 'contracts'):
        flash('Access denied.', 'error')
        return redirect(url_for('index'))
    db = get_db()
    db.execute("UPDATE contract_template SET is_active=0 WHERE id=?", (tid,))
    db.commit()
    flash('Template removed.', 'success')
    return redirect(url_for('contract_templates'))


# ── RFP Tracking ──────────────────────────────────────────────────────────────

RFP_STATUSES = [
    ('sourcing',           'secondary', 'Sourcing'),
    ('proposals_received', 'info',      'Proposals Received'),
    ('negotiating',        'warning',   'Negotiating'),
    ('hotel_selected',     'primary',   'Hotel Selected'),
    ('contracting',        'primary',   'Contracting'),
    ('contracted',         'success',   'Contracted'),
    ('dead',               'danger',    'Dead'),
]
RFP_STATUS_MAP = {s[0]: (s[1], s[2]) for s in RFP_STATUSES}


@app.route('/rfp')
def rfp_dashboard():
    user = get_current_user()
    db = get_db()
    show_archived = request.args.get('archived') == '1'
    acct_filter = get_pickup_account_filter(user)
    base = (1 if show_archived else 0,)
    if acct_filter is None:
        rfps = db.execute(
            'SELECT r.*, (SELECT COUNT(*) FROM rfp_hotel h WHERE h.rfp_id=r.id) AS hotel_count '
            'FROM rfp r WHERE r.archived=? ORDER BY r.created_at DESC', base
        ).fetchall()
    elif acct_filter:
        ph = ','.join('?' * len(acct_filter))
        rfps = db.execute(
            f'SELECT r.*, (SELECT COUNT(*) FROM rfp_hotel h WHERE h.rfp_id=r.id) AS hotel_count '
            f'FROM rfp r WHERE r.archived=? AND r.client_org IN ({ph}) ORDER BY r.created_at DESC',
            base + tuple(acct_filter)
        ).fetchall()
    else:
        rfps = []
    return render_template('rfp_dashboard.html', rfps=rfps, statuses=RFP_STATUS_MAP,
                           show_archived=show_archived, all_statuses=RFP_STATUSES)


@app.route('/rfp/new', methods=['GET', 'POST'])
def rfp_new():
    if request.method == 'POST':
        db = get_db()
        f = request.form
        db.execute('''
            INSERT INTO rfp (rfp_code, client_org, event_name, rfp_name, booking_id,
                start_date, end_date, alt_start_date, alt_end_date,
                alt_start_date_2, alt_end_date_2, alt_start_date_3, alt_end_date_3,
                peak_rooms, total_room_nights, total_attendees, f_and_b_budget,
                response_due_date, decision_due_date, status, notes)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        ''', (
            f.get('rfp_code', '').strip() or None,
            f.get('client_org', '').strip(),
            f.get('event_name', '').strip() or None,
            f.get('rfp_name', '').strip() or None,
            f.get('booking_id', '').strip() or None,
            f.get('start_date') or None,
            f.get('end_date') or None,
            f.get('alt_start_date') or None,
            f.get('alt_end_date') or None,
            f.get('alt_start_date_2') or None,
            f.get('alt_end_date_2') or None,
            f.get('alt_start_date_3') or None,
            f.get('alt_end_date_3') or None,
            f.get('peak_rooms') or None,
            f.get('total_room_nights') or None,
            f.get('total_attendees') or None,
            f.get('f_and_b_budget') or None,
            f.get('response_due_date') or None,
            f.get('decision_due_date') or None,
            f.get('status', 'sourcing'),
            f.get('notes', '').strip() or None,
        ))
        db.commit()
        new_id = db.execute('SELECT last_insert_rowid()').fetchone()[0]
        # Save uploaded RFP document if provided
        rfp_file = request.files.get('rfp_file')
        if rfp_file and rfp_file.filename:
            db.execute("UPDATE rfp SET rfp_filename=?, rfp_data=? WHERE id=?",
                       (rfp_file.filename, rfp_file.read(), new_id))
            db.commit()
        flash('RFP created.', 'success')
        return redirect(url_for('rfp_detail', rid=new_id))
    return render_template('rfp_form.html', rfp=None, all_statuses=RFP_STATUSES)


@app.route('/rfp/<int:rid>')
def rfp_detail(rid):
    user = get_current_user()
    db = get_db()
    rfp = db.execute('SELECT * FROM rfp WHERE id=?', (rid,)).fetchone()
    if not rfp:
        flash('RFP not found.', 'error')
        return redirect(url_for('rfp_dashboard'))
    acct_filter = get_pickup_account_filter(user)
    if acct_filter is not None and rfp['client_org'] not in acct_filter:
        flash('You do not have access to this RFP.', 'error')
        return redirect(url_for('rfp_dashboard'))
    _hotel_order = {'selected':0,'shortlisted':1,'proposal_received':2,'pending':3,'eliminated':4,'declined':5}
    hotels = db.execute('SELECT * FROM rfp_hotel WHERE rfp_id=?', (rid,)).fetchall()
    hotels = sorted(hotels, key=lambda h: _hotel_order.get(h['status'] or 'pending', 3))
    notes = db.execute('SELECT * FROM rfp_note WHERE rfp_id=? ORDER BY note_date DESC', (rid,)).fetchall()
    return render_template('rfp_detail.html', rfp=rfp, hotels=hotels, notes=notes,
                           statuses=RFP_STATUS_MAP, all_statuses=RFP_STATUSES)


@app.route('/rfp/<int:rid>/edit', methods=['GET', 'POST'])
def rfp_edit(rid):
    user = get_current_user()
    db = get_db()
    rfp = db.execute('SELECT * FROM rfp WHERE id=?', (rid,)).fetchone()
    if not rfp:
        flash('RFP not found.', 'error')
        return redirect(url_for('rfp_dashboard'))
    acct_filter = get_pickup_account_filter(user)
    if acct_filter is not None and rfp['client_org'] not in acct_filter:
        flash('You do not have access to this RFP.', 'error')
        return redirect(url_for('rfp_dashboard'))
    if request.method == 'POST':
        f = request.form
        db.execute('''
            UPDATE rfp SET rfp_code=?, client_org=?, event_name=?, rfp_name=?,
                booking_id=?, start_date=?, end_date=?, alt_start_date=?, alt_end_date=?,
                alt_start_date_2=?, alt_end_date_2=?, alt_start_date_3=?, alt_end_date_3=?,
                peak_rooms=?, total_room_nights=?, total_attendees=?, f_and_b_budget=?,
                response_due_date=?, decision_due_date=?, status=?, notes=?,
                updated_at=datetime('now')
            WHERE id=?
        ''', (
            f.get('rfp_code', '').strip() or None,
            f.get('client_org', '').strip(),
            f.get('event_name', '').strip() or None,
            f.get('rfp_name', '').strip() or None,
            f.get('booking_id', '').strip() or None,
            f.get('start_date') or None,
            f.get('end_date') or None,
            f.get('alt_start_date') or None,
            f.get('alt_end_date') or None,
            f.get('alt_start_date_2') or None,
            f.get('alt_end_date_2') or None,
            f.get('alt_start_date_3') or None,
            f.get('alt_end_date_3') or None,
            f.get('peak_rooms') or None,
            f.get('total_room_nights') or None,
            f.get('total_attendees') or None,
            f.get('f_and_b_budget') or None,
            f.get('response_due_date') or None,
            f.get('decision_due_date') or None,
            f.get('status', 'sourcing'),
            f.get('notes', '').strip() or None,
            rid,
        ))
        # Save uploaded RFP document if provided
        rfp_file = request.files.get('rfp_file')
        if rfp_file and rfp_file.filename:
            db.execute("UPDATE rfp SET rfp_filename=?, rfp_data=? WHERE id=?",
                       (rfp_file.filename, rfp_file.read(), rid))
        db.commit()
        flash('RFP updated.', 'success')
        return redirect(url_for('rfp_detail', rid=rid))
    return render_template('rfp_form.html', rfp=rfp, all_statuses=RFP_STATUSES)


@app.route('/rfp/<int:rid>/import-crf', methods=['POST'])
def rfp_import_crf(rid):
    from pickup_utils import parse_crf_excel
    db = get_db()
    rfp = db.execute("SELECT * FROM rfp WHERE id=?", (rid,)).fetchone()
    if not rfp:
        flash('RFP not found.', 'error')
        return redirect(url_for('rfp_dashboard'))
    f = request.files.get('crf_file')
    if not f or not f.filename:
        flash('No file selected.', 'error')
        return redirect(url_for('rfp_detail', rid=rid))
    file_bytes = f.read()
    crf_filename = f.filename
    db.execute("UPDATE rfp SET crf_filename=?, crf_data=?, updated_at=datetime('now') WHERE id=?",
               (crf_filename, file_bytes, rid))
    db.commit()
    try:
        result = parse_crf_excel(file_bytes)
    except Exception as e:
        flash(f'Error parsing CRF: {e}', 'error')
        return redirect(url_for('rfp_detail', rid=rid))
    # Fill blank meta fields from CRF
    meta = result.get('rfp_meta', {})
    updates, vals = [], []
    for field in ('event_name', 'response_due_date', 'decision_due_date'):
        if meta.get(field) and not rfp[field]:
            updates.append(f'{field}=?')
            vals.append(meta[field])
    if updates:
        vals.append(rid)
        db.execute(f"UPDATE rfp SET {', '.join(updates)}, updated_at=datetime('now') WHERE id=?", vals)
        db.commit()
    hotels = result.get('hotels', [])
    n_proposals = sum(1 for h in hotels if h['status'] == 'proposal_received')
    n_declined  = sum(1 for h in hotels if h['status'] == 'declined')
    return render_template('rfp_crf_review.html', rfp=rfp, hotels=hotels,
                           n_proposals=n_proposals, n_declined=n_declined,
                           crf_filename=crf_filename)


@app.route('/rfp/<int:rid>/import-crf/confirm', methods=['POST'])
def rfp_import_crf_confirm(rid):
    from pickup_utils import parse_crf_excel
    db = get_db()
    rfp = db.execute("SELECT * FROM rfp WHERE id=?", (rid,)).fetchone()
    if not rfp:
        flash('RFP not found.', 'error')
        return redirect(url_for('rfp_dashboard'))
    crf_data = rfp['crf_data'] if rfp['crf_data'] else None
    if not crf_data:
        flash('No CRF file stored — please re-upload the CRF.', 'error')
        return redirect(url_for('rfp_detail', rid=rid))
    try:
        hotels = parse_crf_excel(bytes(crf_data)).get('hotels', [])
    except Exception as e:
        flash(f'Error re-parsing CRF: {e}', 'error')
        return redirect(url_for('rfp_detail', rid=rid))

    def _fv(s):
        try: return float(s) if s not in (None,'','None') else None
        except: return None
    def _iv(s):
        try: return int(s) if s not in (None,'','None') else None
        except: return None

    inserted = updated = 0
    any_proposal = False
    _GENERIC = {
        'hotel','hotels','resort','resorts','suites','suite','inn','inns',
        'motel','lodge','lodges','hilton','marriott','hyatt','sheraton',
        'westin','omni','embassy','extended','stay','holiday','best','western',
        'comfort','quality','sleep','springhill','the','and','spa','conference',
        'center','collection','autograph','renaissance','tapestry','curio',
        'home2','homewood','hampton','doubletree','courtyard','fairfield',
        'towneplace','graduate','thompson','kimpton','indigo','vignette','by','at','of',
    }

    for idx, h in enumerate(hotels):
        if request.form.get(f'include_{idx}') != '1':
            continue
        hotel_name = (h.get('hotel_name') or '').strip()
        if not hotel_name:
            continue
        status = h.get('status', 'pending')
        if status == 'proposal_received':
            any_proposal = True
        proposed_rate   = _fv(request.form.get(f'rate_{idx}'))  or _fv(h.get('proposed_rate'))
        f_and_b_minimum = _fv(request.form.get(f'fab_{idx}'))   or _fv(h.get('f_and_b_minimum'))
        comm_raw        = request.form.get(f'comm_{idx}')
        commission_pct  = (_fv(comm_raw) / 100.0 if _fv(comm_raw) else None) or _fv(h.get('commission_pct'))
        notes_val       = request.form.get(f'notes_{idx}','').strip() or h.get('notes') or None
        city  = h.get('city') or None
        state = h.get('state') or None
        crf_row_data = h.get('crf_row_data') or None

        # Match existing hotel by exact name, then word overlap
        existing = db.execute(
            "SELECT id, crf_version FROM rfp_hotel WHERE rfp_id=? AND LOWER(hotel_name)=LOWER(?)",
            (rid, hotel_name)
        ).fetchone()
        if not existing:
            all_hotels = db.execute("SELECT id, hotel_name, crf_version FROM rfp_hotel WHERE rfp_id=?", (rid,)).fetchall()
            h_words = set(w.lower() for w in hotel_name.split() if len(w) > 3 and w.lower() not in _GENERIC)
            for ah in all_hotels:
                ah_words = set(w.lower() for w in ah['hotel_name'].split() if len(w) > 3 and w.lower() not in _GENERIC)
                if h_words and ah_words and (h_words & ah_words):
                    existing = ah
                    break

        new_ver = (existing['crf_version'] or 0) + 1 if existing else 1
        if existing:
            db.execute('''
                UPDATE rfp_hotel SET
                    city=COALESCE(?,city), state=COALESCE(?,state),
                    contact_name=COALESCE(?,contact_name),
                    contact_email=COALESCE(?,contact_email),
                    contact_phone=COALESCE(?,contact_phone),
                    contact_title=COALESCE(?,contact_title),
                    status=?, proposed_rate=?, commission_pct=?,
                    f_and_b_minimum=?, attrition_pct=?, cutoff_days=?,
                    concessions=?, notes=COALESCE(?,notes),
                    crf_row_data=?, crf_version=?, updated_at=datetime('now')
                WHERE id=?
            ''', (city, state,
                  h.get('contact_name') or None, h.get('contact_email') or None,
                  h.get('contact_phone') or None, h.get('contact_title') or None,
                  status, proposed_rate, commission_pct, f_and_b_minimum,
                  _fv(h.get('attrition_pct')), _iv(h.get('cutoff_days')),
                  h.get('concessions') or None, notes_val,
                  crf_row_data, new_ver, existing['id']))
            updated += 1
        else:
            db.execute('''
                INSERT INTO rfp_hotel (rfp_id, hotel_name, city, state,
                    contact_name, contact_email, contact_phone, contact_title,
                    status, proposed_rate, commission_pct, f_and_b_minimum,
                    attrition_pct, cutoff_days, concessions, notes,
                    crf_row_data, crf_version, updated_at)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,datetime('now'))
            ''', (rid, hotel_name, city, state,
                  h.get('contact_name') or None, h.get('contact_email') or None,
                  h.get('contact_phone') or None, h.get('contact_title') or None,
                  status, proposed_rate, commission_pct, f_and_b_minimum,
                  _fv(h.get('attrition_pct')), _iv(h.get('cutoff_days')),
                  h.get('concessions') or None, notes_val,
                  crf_row_data, new_ver))
            inserted += 1

    if any_proposal and rfp['status'] == 'sourcing':
        db.execute("UPDATE rfp SET status='proposals_received', updated_at=datetime('now') WHERE id=?", (rid,))
    db.commit()
    parts = []
    if inserted: parts.append(f'{inserted} hotel{"s" if inserted!=1 else ""} added')
    if updated:  parts.append(f'{updated} hotel{"s" if updated!=1 else ""} updated')
    flash((', '.join(parts) + '.') if parts else 'No changes made.', 'success' if parts else 'info')
    return redirect(url_for('rfp_detail', rid=rid))


@app.route('/rfp/parse-document', methods=['POST'])
def rfp_parse_document():
    """AJAX — parse an uploaded Cvent RFP .docx and return field values as JSON."""
    from pickup_utils import parse_rfp_docx
    f = request.files.get('rfp_file')
    if not f or not f.filename.lower().endswith('.docx'):
        return jsonify({})
    try:
        return jsonify(parse_rfp_docx(f.read()))
    except Exception as e:
        return jsonify({'_error': str(e)})


@app.route('/rfp/<int:rid>/archive', methods=['POST'])
def rfp_archive(rid):
    db = get_db()
    db.execute("UPDATE rfp SET archived=1 WHERE id=?", (rid,))
    db.commit()
    flash('RFP archived.', 'success')
    return redirect(url_for('rfp_dashboard'))


@app.route('/rfp/<int:rid>/unarchive', methods=['POST'])
def rfp_unarchive(rid):
    db = get_db()
    db.execute("UPDATE rfp SET archived=0 WHERE id=?", (rid,))
    db.commit()
    flash('RFP restored.', 'success')
    return redirect(url_for('rfp_dashboard', archived=1))


@app.route('/rfp/<int:rid>/hotel/add', methods=['GET', 'POST'])
def rfp_hotel_add(rid):
    db = get_db()
    rfp = db.execute('SELECT * FROM rfp WHERE id=?', (rid,)).fetchone()
    if not rfp:
        flash('RFP not found.', 'error')
        return redirect(url_for('rfp_dashboard'))
    if request.method == 'POST':
        db.execute('''
            INSERT INTO rfp_hotel (rfp_id, hotel_name, brand, city, state,
                contact_name, contact_email, contact_phone, contact_title,
                proposed_rate, commission_pct, f_and_b_minimum, meeting_room_rental,
                attrition_pct, cutoff_days, concessions, notes)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        ''', (
            rid,
            request.form.get('hotel_name', '').strip(),
            request.form.get('brand', '').strip() or None,
            request.form.get('city', '').strip() or None,
            request.form.get('state', '').strip() or None,
            request.form.get('contact_name', '').strip() or None,
            request.form.get('contact_email', '').strip() or None,
            request.form.get('contact_phone', '').strip() or None,
            request.form.get('contact_title', '').strip() or None,
            request.form.get('proposed_rate') or None,
            request.form.get('commission_pct') or None,
            request.form.get('f_and_b_minimum') or None,
            request.form.get('meeting_room_rental') or None,
            request.form.get('attrition_pct') or None,
            request.form.get('cutoff_days') or None,
            request.form.get('concessions', '').strip() or None,
            request.form.get('notes', '').strip() or None,
        ))
        db.commit()
        flash('Hotel added.', 'success')
        return redirect(url_for('rfp_detail', rid=rid))
    return render_template('rfp_hotel_form.html', rfp=rfp, hotel=None)


@app.route('/rfp/<int:rid>/hotel/<int:hid>/edit', methods=['GET', 'POST'])
def rfp_hotel_edit(rid, hid):
    db = get_db()
    rfp = db.execute('SELECT * FROM rfp WHERE id=?', (rid,)).fetchone()
    hotel = db.execute('SELECT * FROM rfp_hotel WHERE id=? AND rfp_id=?', (hid, rid)).fetchone()
    if not rfp or not hotel:
        flash('Not found.', 'error')
        return redirect(url_for('rfp_dashboard'))
    if request.method == 'POST':
        db.execute('''
            UPDATE rfp_hotel SET hotel_name=?, brand=?, city=?, state=?,
                contact_name=?, contact_email=?, contact_phone=?, contact_title=?,
                proposed_rate=?, commission_pct=?, f_and_b_minimum=?,
                meeting_room_rental=?, attrition_pct=?, cutoff_days=?,
                concessions=?, notes=?, updated_at=datetime('now')
            WHERE id=? AND rfp_id=?
        ''', (
            request.form.get('hotel_name', '').strip(),
            request.form.get('brand', '').strip() or None,
            request.form.get('city', '').strip() or None,
            request.form.get('state', '').strip() or None,
            request.form.get('contact_name', '').strip() or None,
            request.form.get('contact_email', '').strip() or None,
            request.form.get('contact_phone', '').strip() or None,
            request.form.get('contact_title', '').strip() or None,
            request.form.get('proposed_rate') or None,
            request.form.get('commission_pct') or None,
            request.form.get('f_and_b_minimum') or None,
            request.form.get('meeting_room_rental') or None,
            request.form.get('attrition_pct') or None,
            request.form.get('cutoff_days') or None,
            request.form.get('concessions', '').strip() or None,
            request.form.get('notes', '').strip() or None,
            hid, rid,
        ))
        db.commit()
        flash('Hotel updated.', 'success')
        return redirect(url_for('rfp_detail', rid=rid))
    return render_template('rfp_hotel_form.html', rfp=rfp, hotel=hotel)


@app.route('/rfp/<int:rid>/hotel/<int:hid>/delete', methods=['POST'])
def rfp_hotel_delete(rid, hid):
    db = get_db()
    db.execute('DELETE FROM rfp_hotel WHERE id=? AND rfp_id=?', (hid, rid))
    db.commit()
    flash('Hotel removed.', 'success')
    return redirect(url_for('rfp_detail', rid=rid))


@app.route('/rfp/<int:rid>/hotel/<int:hid>/select', methods=['POST'])
def rfp_hotel_select(rid, hid):
    db = get_db()
    # Select this hotel and update the RFP status
    db.execute("UPDATE rfp_hotel SET status='selected' WHERE id=? AND rfp_id=?", (hid, rid))
    db.execute("UPDATE rfp SET status='hotel_selected', updated_at=datetime('now') WHERE id=?", (rid,))
    # Handle other hotels: remove or mark eliminated
    others_action = request.form.get('others_action', 'keep')
    if others_action == 'remove':
        db.execute("DELETE FROM rfp_hotel WHERE rfp_id=? AND id!=?", (rid, hid))
    else:
        db.execute("UPDATE rfp_hotel SET status='eliminated' WHERE rfp_id=? AND id!=? AND status NOT IN ('selected','declined')", (rid, hid))
    # Optional proposal upload
    pf = request.files.get('proposal_file')
    if pf and pf.filename:
        db.execute(
            "UPDATE rfp_hotel SET proposal_filename=?, proposal_data=?, updated_at=datetime('now') WHERE id=? AND rfp_id=?",
            (pf.filename, pf.read(), hid, rid)
        )
    db.commit()
    flash('Hotel selected.', 'success')
    return redirect(url_for('rfp_detail', rid=rid))


@app.route('/rfp/<int:rid>/hotel/<int:hid>/upload-proposal', methods=['POST'])
def rfp_hotel_upload_proposal(rid, hid):
    f = request.files.get('proposal_file')
    if not f or not f.filename:
        flash('No file selected.', 'error')
        return redirect(url_for('rfp_detail', rid=rid))
    db = get_db()
    db.execute(
        "UPDATE rfp_hotel SET proposal_filename=?, proposal_data=?, updated_at=datetime('now') WHERE id=? AND rfp_id=?",
        (f.filename, f.read(), hid, rid)
    )
    db.commit()
    flash(f'Proposal "{f.filename}" uploaded.', 'success')
    return redirect(url_for('rfp_detail', rid=rid))


@app.route('/rfp/<int:rid>/hotel/<int:hid>/proposal/download')
def rfp_hotel_proposal_download(rid, hid):
    import io as _io
    db = get_db()
    h = db.execute('SELECT * FROM rfp_hotel WHERE id=? AND rfp_id=?', (hid, rid)).fetchone()
    if not h or not h['proposal_data']:
        flash('Proposal not found.', 'error')
        return redirect(url_for('rfp_detail', rid=rid))
    return send_file(_io.BytesIO(h['proposal_data']),
                     download_name=h['proposal_filename'] or 'proposal.pdf',
                     as_attachment=True)


@app.route('/rfp/<int:rid>/upload-contract', methods=['POST'])
def rfp_upload_contract(rid):
    """Step 1 — AI-parse an uploaded hotel contract and show a review screen."""
    import uuid, tempfile
    from pickup_utils import parse_contract_document
    db = get_db()
    rfp = db.execute('SELECT * FROM rfp WHERE id=?', (rid,)).fetchone()
    if not rfp:
        flash('RFP not found.', 'error')
        return redirect(url_for('rfp_dashboard'))
    # Save booking_id to rfp record if provided (required by modal)
    booking_id = request.form.get('booking_id', '').strip()
    if not booking_id:
        flash('Booking ID is required before importing a contract.', 'error')
        return redirect(url_for('rfp_detail', rid=rid))
    if booking_id != (rfp['booking_id'] or '').strip():
        db.execute("UPDATE rfp SET booking_id=?, updated_at=datetime('now') WHERE id=?",
                   (booking_id, rid))
        db.commit()
        rfp = db.execute('SELECT * FROM rfp WHERE id=?', (rid,)).fetchone()

    f = request.files.get('contract_file')
    if not f or not f.filename:
        flash('No file selected.', 'error')
        return redirect(url_for('rfp_detail', rid=rid))
    file_bytes = f.read()

    # Find the selected (or best available) hotel to pre-fill contacts
    selected_hotel = db.execute(
        "SELECT * FROM rfp_hotel WHERE rfp_id=? AND status='selected' "
        "ORDER BY updated_at DESC LIMIT 1", (rid,)
    ).fetchone()
    if not selected_hotel:
        selected_hotel = db.execute(
            "SELECT * FROM rfp_hotel WHERE rfp_id=? "
            "AND status NOT IN ('declined','eliminated') "
            "ORDER BY updated_at DESC LIMIT 1", (rid,)
        ).fetchone()

    extracted = parse_contract_document(file_bytes, filename=f.filename)
    if extracted.get('error'):
        flash(f'Could not parse contract: {extracted["error"]}', 'error')
        return redirect(url_for('rfp_detail', rid=rid))

    # Save to temp file so confirm step doesn't need a re-upload
    tmp_id   = str(uuid.uuid4())
    tmp_path = os.path.join(tempfile.gettempdir(), f'rfp_contract_{tmp_id}')
    with open(tmp_path, 'wb') as fh:
        fh.write(file_bytes)

    return render_template('rfp_contract_review.html',
                           rfp=rfp,
                           selected_hotel=selected_hotel,
                           extracted=extracted,
                           filename=f.filename,
                           tmp_id=tmp_id)


@app.route('/rfp/<int:rid>/upload-contract/confirm', methods=['POST'])
def rfp_upload_contract_confirm(rid):
    """Step 2 — save AI-extracted contract data to rfp, hotel, and pickup_config."""
    import tempfile
    from datetime import date as _date
    db = get_db()
    rfp = db.execute('SELECT * FROM rfp WHERE id=?', (rid,)).fetchone()
    if not rfp:
        flash('RFP not found.', 'error')
        return redirect(url_for('rfp_dashboard'))

    # ── Parse form ────────────────────────────────────────────────────────────
    raw_block    = request.form.get('contracted_block', '{}')
    block        = json.loads(raw_block)
    rate_str     = request.form.get('contracted_rate', '').strip()
    cutoff_str   = request.form.get('cutoff_date', '').strip()
    atr_str      = request.form.get('attrition_pct', '').strip()
    hotel_contact        = request.form.get('hotel_contact', '').strip() or None
    hotel_contact_email  = request.form.get('hotel_contact_email', '').strip() or None
    hotel_contact2       = request.form.get('hotel_contact2', '').strip() or hotel_contact
    hotel_contact2_email = request.form.get('hotel_contact2_email', '').strip() or hotel_contact_email
    group_contact        = request.form.get('group_contact', '').strip() or None
    group_contact_email  = request.form.get('group_contact_email', '').strip() or None

    contracted_rate = float(rate_str) if rate_str else None
    attrition_pct   = float(atr_str)  if atr_str  else None

    # ── Derive rfp-level fields from block ────────────────────────────────────
    dates             = sorted(block.keys())
    start_date        = dates[0]  if dates else None
    end_date          = dates[-1] if dates else None
    peak_rooms        = max(block.values()) if block else None
    total_room_nights = sum(block.values()) if block else None

    cutoff_days = None
    if cutoff_str and start_date:
        try:
            cutoff_days = (_date.fromisoformat(start_date) -
                           _date.fromisoformat(cutoff_str)).days
        except Exception:
            pass

    # ── File blob from temp file written in step 1 ────────────────────────────
    contract_filename = request.form.get('_contract_filename', '')
    tmp_id   = request.form.get('_tmp_id', '')
    tmp_path = os.path.join(tempfile.gettempdir(), f'rfp_contract_{tmp_id}') if tmp_id else ''
    file_blob = None
    if tmp_path and os.path.exists(tmp_path):
        with open(tmp_path, 'rb') as fh:
            file_blob = fh.read()
        try:
            os.remove(tmp_path)
        except Exception:
            pass

    # ── Update rfp record ─────────────────────────────────────────────────────
    db.execute('''
        UPDATE rfp SET
            start_date        = COALESCE(?, start_date),
            end_date          = COALESCE(?, end_date),
            peak_rooms        = COALESCE(?, peak_rooms),
            total_room_nights = COALESCE(?, total_room_nights),
            contract_filename = ?,
            contract_data     = ?,
            status            = CASE WHEN status NOT IN ('contracted','dead')
                                     THEN 'contracted' ELSE status END,
            updated_at        = datetime('now')
        WHERE id=?
    ''', (start_date, end_date, peak_rooms, total_room_nights,
          contract_filename or None, file_blob, rid))

    # ── Update selected hotel record ──────────────────────────────────────────
    selected_hotel = db.execute(
        "SELECT * FROM rfp_hotel WHERE rfp_id=? AND status='selected' "
        "ORDER BY updated_at DESC LIMIT 1", (rid,)
    ).fetchone()
    if not selected_hotel:
        selected_hotel = db.execute(
            "SELECT * FROM rfp_hotel WHERE rfp_id=? "
            "AND status NOT IN ('declined','eliminated') "
            "ORDER BY updated_at DESC LIMIT 1", (rid,)
        ).fetchone()
    if selected_hotel:
        db.execute('''
            UPDATE rfp_hotel SET
                proposed_rate = COALESCE(?, proposed_rate),
                attrition_pct = COALESCE(?, attrition_pct),
                cutoff_days   = COALESCE(?, cutoff_days),
                contact_name  = COALESCE(?, contact_name),
                contact_email = COALESCE(?, contact_email),
                updated_at    = datetime('now')
            WHERE id=?
        ''', (contracted_rate, attrition_pct, cutoff_days,
              hotel_contact, hotel_contact_email, selected_hotel['id']))

    # ── Update or create pickup_config ────────────────────────────────────────
    pickup_updated = False
    if block and rfp['booking_id']:
        existing_pc = db.execute(
            'SELECT id FROM pickup_config WHERE CAST(booking_id AS INTEGER)=CAST(? AS INTEGER)',
            (rfp['booking_id'],)
        ).fetchone()
        hotel_name = selected_hotel['hotel_name'] if selected_hotel else (rfp['event_name'] or '')
        org        = rfp['client_org'] or ''
        event_name = rfp['event_name'] or ''
        cutoff_iso = cutoff_str or None

        if existing_pc:
            db.execute('''
                UPDATE pickup_config SET
                    contracted_block     = ?,
                    contracted_rate      = COALESCE(?, contracted_rate),
                    cutoff_date          = COALESCE(?, cutoff_date),
                    attrition_pct        = COALESCE(?, attrition_pct),
                    block_is_estimated   = 0,
                    contract_filename    = COALESCE(?, contract_filename),
                    contract_data        = COALESCE(?, contract_data),
                    hotel_contact        = COALESCE(?, hotel_contact),
                    hotel_contact_email  = COALESCE(?, hotel_contact_email),
                    hotel_contact2       = COALESCE(?, hotel_contact2),
                    hotel_contact2_email = COALESCE(?, hotel_contact2_email),
                    group_contact        = COALESCE(?, group_contact),
                    group_contact_email  = COALESCE(?, group_contact_email)
                WHERE id=?
            ''', (json.dumps(block), contracted_rate, cutoff_iso, attrition_pct,
                  contract_filename or None, file_blob,
                  hotel_contact, hotel_contact_email,
                  hotel_contact2, hotel_contact2_email,
                  group_contact, group_contact_email,
                  existing_pc['id']))
            pickup_updated = True
        else:
            from datetime import timedelta as _td
            cutoff_default = (
                cutoff_iso if cutoff_iso else
                ((_date.fromisoformat(start_date) - _td(days=30)).isoformat()
                 if start_date else None)
            )
            db.execute('''
                INSERT INTO pickup_config
                    (booking_id, organization, event_name, hotel,
                     contracted_block, contracted_rate, cutoff_date,
                     attrition_pct, status, block_is_estimated,
                     contract_filename, contract_data,
                     hotel_contact, hotel_contact_email,
                     hotel_contact2, hotel_contact2_email,
                     group_contact, group_contact_email,
                     hotel_contacts, cc_emails, force_current, force_past)
                VALUES (?,?,?,?,?,?,?,?,'active',0,?,?,?,?,?,?,?,?,'[]','[]',0,0)
            ''', (str(rfp['booking_id']), org, event_name, hotel_name,
                  json.dumps(block), contracted_rate, cutoff_default,
                  attrition_pct or 0.80,
                  contract_filename or None, file_blob,
                  hotel_contact, hotel_contact_email,
                  hotel_contact2, hotel_contact2_email,
                  group_contact, group_contact_email))
            pickup_updated = True

    db.commit()
    msg = 'Contract saved — RFP and hotel updated.'
    if pickup_updated:
        msg += ' Pickup tracking updated with real block data.'
    flash(msg, 'success')
    return redirect(url_for('rfp_detail', rid=rid))


@app.route('/rfp/<int:rid>/contract/download')
def rfp_contract_download(rid):
    import io as _io
    db = get_db()
    rfp = db.execute('SELECT contract_filename, contract_data FROM rfp WHERE id=?', (rid,)).fetchone()
    if not rfp or not rfp['contract_data']:
        flash('Contract not found.', 'error')
        return redirect(url_for('rfp_detail', rid=rid))
    return send_file(_io.BytesIO(rfp['contract_data']),
                     download_name=rfp['contract_filename'] or 'contract.pdf',
                     as_attachment=True)


@app.route('/rfp/<int:rid>/upload-rfp', methods=['POST'])
def rfp_upload_document(rid):
    f = request.files.get('rfp_file')
    if not f or not f.filename:
        flash('No file selected.', 'error')
        return redirect(url_for('rfp_detail', rid=rid))
    db = get_db()
    db.execute("UPDATE rfp SET rfp_filename=?, rfp_data=?, updated_at=datetime('now') WHERE id=?",
               (f.filename, f.read(), rid))
    db.commit()
    flash(f'RFP document "{f.filename}" uploaded.', 'success')
    return redirect(url_for('rfp_detail', rid=rid))


@app.route('/rfp/<int:rid>/rfp/download')
def rfp_document_download(rid):
    import io as _io
    db = get_db()
    rfp = db.execute('SELECT * FROM rfp WHERE id=?', (rid,)).fetchone()
    if not rfp or not rfp['rfp_data']:
        flash('RFP document not found.', 'error')
        return redirect(url_for('rfp_detail', rid=rid))
    return send_file(_io.BytesIO(rfp['rfp_data']),
                     download_name=rfp['rfp_filename'] or 'rfp.pdf',
                     as_attachment=True)


@app.route('/rfp/<int:rid>/note/add', methods=['POST'])
def rfp_note_add(rid):
    db = get_db()
    note_text = request.form.get('note_text', '').strip()
    if note_text:
        db.execute(
            "INSERT INTO rfp_note (rfp_id, note_date, note_type, note_text) VALUES (?,date('now'),?,?)",
            (rid, request.form.get('note_type', 'internal'), note_text)
        )
        db.commit()
    return redirect(url_for('rfp_detail', rid=rid))


@app.route('/rfp/<int:rid>/note/<int:nid>/delete', methods=['POST'])
def rfp_note_delete(rid, nid):
    db = get_db()
    db.execute('DELETE FROM rfp_note WHERE id=? AND rfp_id=?', (nid, rid))
    db.commit()
    return redirect(url_for('rfp_detail', rid=rid))


# ── AI Assistant ──────────────────────────────────────────────────────────────

try:
    import anthropic as _anthropic
    from config import ANTHROPIC_API_KEY as _ANTHROPIC_KEY
    _ANTHROPIC_AVAILABLE = bool(_ANTHROPIC_KEY and _ANTHROPIC_KEY.strip())
except Exception:
    _anthropic = None
    _ANTHROPIC_KEY = ''
    _ANTHROPIC_AVAILABLE = False

_ASSISTANT_SYSTEM = """You are Kristin's personal business assistant for CPAinc, a conference planning and booking company.
You have access to the booking database (SQLite).

Your tools let you:
- query_database: Run read-only SQL against the bookings database

Key tables:
- ReportPipeline: BookingId, BookingName, EventName, AccountName, Customer (hotel), BookingAssociate,
  StartDate, EndDate, BookingStatus, Revenue, USDRevenue, CommissionPercent, USDCommissionableAmount,
  PeakRooms, TotalRoomNights, RoomRate, ContractedAmount, City, State, Country
- ChkRegNote: ChkRegID, BookingID, FinalPayment, Check_, DateOnCheck, DepositDate, EntryDate, Cancelled
- Pickup: rowid, BookingID, Brand, ActualPickup, TotalRevenue, EntryDate
- pickup_config: id, booking_id, organization, event_name, hotel, contracted_block, contracted_rate, status, cutoff_date

Today's date is {today}.

Be concise and helpful. Format numbers as currency where appropriate.
"""

_ASSISTANT_TOOLS = [
    {
        'name': 'query_database',
        'description': 'Run a read-only SQL query against the CPAinc SQLite database.',
        'input_schema': {
            'type': 'object',
            'properties': {
                'sql': {'type': 'string', 'description': 'A valid SQLite SELECT statement'},
            },
            'required': ['sql'],
        },
    },
]


def _run_assistant_tool(tool_name, tool_input):
    try:
        if tool_name == 'query_database':
            sql = tool_input.get('sql', '')
            if any(kw in sql.upper() for kw in ('INSERT', 'UPDATE', 'DELETE', 'DROP', 'ALTER', 'CREATE')):
                return 'Error: Only SELECT queries are allowed.'
            db = get_db()
            rows = db.execute(sql).fetchall()
            if not rows:
                return 'No results.'
            cols = rows[0].keys()
            lines = ['\t'.join(str(c) for c in cols)]
            for r in rows[:100]:
                lines.append('\t'.join(str(r[c]) if r[c] is not None else '' for c in cols))
            return '\n'.join(lines)
    except Exception as e:
        return f'Error: {e}'


@app.route('/assistant', methods=['GET', 'POST'])
def assistant():
    if not _ANTHROPIC_AVAILABLE:
        flash('AI Assistant is not configured (no API key).', 'error')
        return redirect(url_for('index'))
    if request.method == 'GET':
        return render_template('assistant.html')

    data = request.get_json()
    messages = data.get('messages', [])

    client = _anthropic.Anthropic(api_key=_ANTHROPIC_KEY)
    system_prompt = _ASSISTANT_SYSTEM.format(today=datetime.now().strftime('%B %d, %Y'))
    api_messages = [{'role': m['role'], 'content': m['content']} for m in messages]

    max_iterations = 8
    for _ in range(max_iterations):
        response = client.messages.create(
            model='claude-sonnet-4-6',
            max_tokens=4096,
            system=system_prompt,
            tools=_ASSISTANT_TOOLS,
            messages=api_messages,
        )
        assistant_content = response.content
        api_messages.append({'role': 'assistant', 'content': assistant_content})
        if response.stop_reason == 'tool_use':
            tool_results = []
            for block in assistant_content:
                if block.type == 'tool_use':
                    result = _run_assistant_tool(block.name, block.input)
                    tool_results.append({'type': 'tool_result', 'tool_use_id': block.id, 'content': result})
            api_messages.append({'role': 'user', 'content': tool_results})
        else:
            break

    final_text = ''
    for block in response.content:
        if hasattr(block, 'text'):
            final_text += block.text

    return jsonify({'reply': final_text})


# ── Outlook Integration ───────────────────────────────────────────────────────

def _outlook_redirect_uri():
    return request.host_url.rstrip('/') + '/outlook/callback'


@app.route('/outlook/status')
def outlook_status():
    if not _OUTLOOK_AVAILABLE:
        flash('Outlook connector not available.', 'error')
        return redirect(url_for('index'))
    connected = _oc.is_connected()
    user_info = None
    if connected:
        try:
            user_info = _oc.get_user_info()
        except Exception:
            pass
    return render_template('outlook_status.html', connected=connected, user_info=user_info)


@app.route('/outlook/auth')
def outlook_auth():
    if not _OUTLOOK_AVAILABLE:
        return redirect(url_for('index'))
    import urllib.parse
    try:
        from config import MS_CLIENT_ID, MS_TENANT_ID
        scopes = 'Mail.ReadWrite Mail.Send Calendars.ReadWrite Contacts.ReadWrite User.Read'
        params = {
            'client_id': MS_CLIENT_ID,
            'response_type': 'code',
            'redirect_uri': _outlook_redirect_uri(),
            'scope': scopes,
            'response_mode': 'query',
        }
        auth_url = (f'https://login.microsoftonline.com/{MS_TENANT_ID}/oauth2/v2.0/authorize?'
                    + urllib.parse.urlencode(params))
        return redirect(auth_url)
    except Exception as e:
        flash(f'Could not build login URL: {e}', 'error')
        return redirect(url_for('outlook_status'))


@app.route('/outlook/callback')
def outlook_callback():
    if not _OUTLOOK_AVAILABLE:
        return redirect(url_for('index'))
    code = request.args.get('code')
    error = request.args.get('error_description') or request.args.get('error')
    if error:
        flash(f'Microsoft login failed: {error}', 'error')
        return redirect(url_for('outlook_status'))
    if not code:
        flash('No authorization code received.', 'error')
        return redirect(url_for('outlook_status'))
    ok, err = _oc.exchange_code_for_token(code, _outlook_redirect_uri())
    if ok:
        flash('Connected to Microsoft 365 successfully!', 'success')
    else:
        flash(f'Token exchange failed: {err}', 'error')
    return redirect(url_for('outlook_status'))


@app.route('/outlook/disconnect', methods=['POST'])
def outlook_disconnect():
    if _OUTLOOK_AVAILABLE:
        _oc.disconnect()
    flash('Disconnected from Microsoft 365.', 'success')
    return redirect(url_for('outlook_status'))


@app.route('/outlook/inbox')
def outlook_inbox():
    if not _OUTLOOK_AVAILABLE or not _oc.is_connected():
        flash('Connect to Microsoft 365 first.', 'error')
        return redirect(url_for('outlook_status'))
    search = request.args.get('search', '')
    try:
        emails = _oc.get_emails(count=50, search=search if search else None)
    except Exception as e:
        flash(f'Could not load inbox: {e}', 'error')
        emails = []
    return render_template('outlook_inbox.html', emails=emails, search=search)


@app.route('/outlook/calendar')
def outlook_calendar():
    if not _OUTLOOK_AVAILABLE or not _oc.is_connected():
        flash('Connect to Microsoft 365 first.', 'error')
        return redirect(url_for('outlook_status'))
    try:
        events = _oc.get_calendar_events(days_ahead=60, days_back=7)
    except Exception as e:
        flash(f'Could not load calendar: {e}', 'error')
        events = []
    return render_template('outlook_calendar.html', events=events)


# ── Per-user Microsoft OAuth ──────────────────────────────────────────────────

def _ms_redirect_uri():
    base = request.host_url.rstrip('/')
    # Railway is always https — fix scheme if Flask reports http
    if 'railway.app' in base:
        base = base.replace('http://', 'https://')
    return base + '/auth/microsoft/callback'


@app.route('/auth/microsoft')
def auth_microsoft():
    user = get_current_user()
    if not user:
        return redirect(url_for('login'))
    import urllib.parse, secrets as _secrets
    from config import MS_CLIENT_ID, MS_TENANT_ID
    state = _secrets.token_urlsafe(16)
    session['ms_oauth_state'] = state
    params = {
        'client_id':     MS_CLIENT_ID,
        'response_type': 'code',
        'redirect_uri':  _ms_redirect_uri(),
        'scope':         'Mail.ReadWrite User.Read offline_access',
        'response_mode': 'query',
        'state':         state,
        'login_hint':    user['email'],
    }
    url = (f'https://login.microsoftonline.com/{MS_TENANT_ID}/oauth2/v2.0/authorize?'
           + urllib.parse.urlencode(params))
    return redirect(url)


@app.route('/auth/microsoft/callback')
def auth_microsoft_callback():
    user = get_current_user()
    if not user:
        return redirect(url_for('login'))

    error = request.args.get('error_description') or request.args.get('error')
    if error:
        flash(f'Microsoft login failed: {error}', 'error')
        return redirect(url_for('profile'))

    if request.args.get('state') != session.pop('ms_oauth_state', None):
        flash('Invalid state — please try again.', 'error')
        return redirect(url_for('profile'))

    code = request.args.get('code')
    if not code:
        flash('No authorisation code received from Microsoft.', 'error')
        return redirect(url_for('profile'))

    import time, requests as _req
    from config import MS_CLIENT_ID, MS_TENANT_ID, MS_CLIENT_SECRET

    resp = _req.post(
        f'https://login.microsoftonline.com/{MS_TENANT_ID}/oauth2/v2.0/token',
        data={
            'grant_type':    'authorization_code',
            'client_id':     MS_CLIENT_ID,
            'client_secret': MS_CLIENT_SECRET,
            'code':          code,
            'redirect_uri':  _ms_redirect_uri(),
            'scope':         'Mail.ReadWrite User.Read offline_access',
        }
    )
    if resp.status_code != 200:
        flash(f'Token exchange failed: {resp.text[:200]}', 'error')
        return redirect(url_for('profile'))

    tok           = resp.json()
    access_token  = tok['access_token']
    refresh_token = tok.get('refresh_token', '')
    expires_at    = time.time() + tok.get('expires_in', 3600) - 60

    me = _req.get('https://graph.microsoft.com/v1.0/me',
                  headers={'Authorization': f'Bearer {access_token}'},
                  params={'$select': 'mail,userPrincipalName'})
    ms_email = ''
    if me.status_code == 200:
        d = me.json()
        ms_email = d.get('mail') or d.get('userPrincipalName', '')

    db = get_db()
    db.execute('''
        INSERT INTO UserMicrosoftTokens (user_id, access_token, refresh_token, expires_at, ms_user_email)
        VALUES (?, ?, ?, ?, ?)
        ON CONFLICT(user_id) DO UPDATE SET
            access_token  = excluded.access_token,
            refresh_token = excluded.refresh_token,
            expires_at    = excluded.expires_at,
            ms_user_email = excluded.ms_user_email,
            connected_at  = datetime('now')
    ''', (user['id'], access_token, refresh_token, expires_at, ms_email))
    db.commit()

    flash(f'Microsoft account connected ({ms_email}).', 'success')
    return redirect(url_for('profile'))


@app.route('/auth/microsoft/disconnect', methods=['POST'])
def auth_microsoft_disconnect():
    user = get_current_user()
    if not user:
        return redirect(url_for('login'))
    db = get_db()
    db.execute('DELETE FROM UserMicrosoftTokens WHERE user_id = ?', (user['id'],))
    db.commit()
    flash('Microsoft account disconnected.', 'success')
    return redirect(url_for('profile'))


@app.route('/profile')
def profile():
    user = get_current_user()
    if not user:
        return redirect(url_for('login'))
    db = get_db()
    ms_row = db.execute(
        'SELECT ms_user_email, connected_at FROM UserMicrosoftTokens WHERE user_id = ?',
        (user['id'],)
    ).fetchone()
    return render_template('profile.html', user=user, ms_token=ms_row)



@app.route('/admin/change-log')
def admin_change_log():
    user = get_current_user()
    if not user or user['role'] != 'admin':
        flash('Admin access required.', 'error')
        return redirect(url_for('index'))
    db = get_db()
    filter_cid  = request.args.get('cid', '').strip()
    filter_user = request.args.get('user', '').strip()
    filter_date = request.args.get('date', '').strip()
    page     = max(1, int(request.args.get('page', 1)))
    per_page = 100
    where, params = [], []
    if filter_cid:
        where.append('cl.config_id = ?'); params.append(filter_cid)
    if filter_user:
        where.append('cl.username = ?'); params.append(filter_user)
    if filter_date:
        where.append("DATE(cl.timestamp) = ?"); params.append(filter_date)
    where_sql = ('WHERE ' + ' AND '.join(where)) if where else ''
    total = db.execute(f'SELECT COUNT(*) FROM pickup_change_log cl {where_sql}', params).fetchone()[0]
    logs  = db.execute(
        f'''SELECT cl.*, pc.event_name, pc.hotel
            FROM pickup_change_log cl
            LEFT JOIN pickup_config pc ON pc.id = cl.config_id
            {where_sql}
            ORDER BY cl.id DESC LIMIT ? OFFSET ?''',
        params + [per_page, (page-1)*per_page]
    ).fetchall()
    all_users = [r[0] for r in db.execute(
        "SELECT DISTINCT username FROM pickup_change_log WHERE username IS NOT NULL ORDER BY username"
    ).fetchall()]
    total_pages = max(1, (total + per_page - 1) // per_page)
    return render_template('change_log.html', logs=logs, total=total, page=page,
                           total_pages=total_pages, per_page=per_page,
                           filter_cid=filter_cid, filter_user=filter_user,
                           filter_date=filter_date, all_users=all_users)


@app.route('/admin/run-daily-tasks')
def admin_run_daily_tasks():
    """Called by an external cron service daily. Protected by DAILY_TASK_KEY.
    Sends alert emails for overdue items and a DB backup every Sunday."""
    key = request.args.get('key', '')
    if DAILY_TASK_KEY and key != DAILY_TASK_KEY:
        return 'Forbidden', 403
    if not MAIL_TO:
        return jsonify({'ok': False, 'error': 'MAIL_TO not configured'}), 500

    db      = get_db()
    today   = datetime.today().strftime('%Y-%m-%d')
    log     = []

    # ── 1. Build alert summary ────────────────────────────────────────────────
    configs = db.execute(
        "SELECT * FROM pickup_config WHERE status='active'"
    ).fetchall()
    has_hhr = {str(r[0]) for r in db.execute(
        "SELECT DISTINCT booking_id FROM housing_history_files WHERE booking_id IS NOT NULL"
    ).fetchall()}
    has_history = {r[0] for r in db.execute(
        "SELECT DISTINCT config_id FROM pickup_weekly"
    ).fetchall()}
    latest_pickup = {r[0]: r[1] for r in db.execute(
        "SELECT config_id, MAX(report_date) FROM pickup_weekly GROUP BY config_id"
    ).fetchall()}

    alerts = {'overdue': [], 'cutoff_soon': [], 'ended_no_hhr': []}

    for cfg in configs:
        cid   = cfg['id']
        block = {}
        try:
            block = json.loads(cfg['contracted_block'] or '{}')
        except Exception:
            pass
        dates       = sorted(block.keys())
        event_start = dates[0]  if dates else None
        event_end   = dates[-1] if dates else None
        cutoff      = (cfg['cutoff_date'] or '').strip()
        is_started  = event_start and event_start <= today
        is_ended    = event_end   and event_end   <  today
        is_current  = is_started and not is_ended

        if is_current:
            last_rpt = latest_pickup.get(cid)
            if not last_rpt:
                alerts['overdue'].append(f"  • {cfg['event_name']} @ {cfg['hotel']} — no reports yet")
            else:
                days_since = (datetime.strptime(today, '%Y-%m-%d') -
                              datetime.strptime(last_rpt, '%Y-%m-%d')).days
                if days_since >= 7:
                    alerts['overdue'].append(
                        f"  • {cfg['event_name']} @ {cfg['hotel']} — last report {days_since} days ago ({last_rpt})"
                    )
            if cutoff:
                try:
                    days_to = (datetime.strptime(cutoff, '%Y-%m-%d') -
                               datetime.strptime(today, '%Y-%m-%d')).days
                    if 0 <= days_to <= 7:
                        alerts['cutoff_soon'].append(
                            f"  • {cfg['event_name']} @ {cfg['hotel']} — cutoff in {days_to} day(s) ({cutoff})"
                        )
                except Exception:
                    pass

        if is_ended and str(cfg['booking_id'] or '') not in has_hhr:
            alerts['ended_no_hhr'].append(
                f"  • {cfg['event_name']} @ {cfg['hotel']} — ended {event_end}"
            )

    # ── 2. Build email body ───────────────────────────────────────────────────
    sections = []
    if alerts['overdue']:
        sections.append('<h3 style="color:#c0392b">⚠️ Overdue Pickup Reports</h3><pre style="background:#fdf2f2;padding:12px">'
                        + '\n'.join(alerts['overdue']) + '</pre>')
    if alerts['cutoff_soon']:
        sections.append('<h3 style="color:#e67e22">📅 Cutoff Approaching (within 7 days)</h3><pre style="background:#fef9e7;padding:12px">'
                        + '\n'.join(alerts['cutoff_soon']) + '</pre>')
    if alerts['ended_no_hhr']:
        sections.append('<h3 style="color:#8e44ad">📂 Events Ended — No HHR Uploaded</h3><pre style="background:#f5eef8;padding:12px">'
                        + '\n'.join(alerts['ended_no_hhr']) + '</pre>')

    sent_alert = False
    if sections:
        body = (f'<p>CPAinc daily alert — {today}</p>'
                + ''.join(sections)
                + '<p style="color:#888;font-size:.85em">Sent automatically by CPAinc. '
                  'View the <a href="https://cpainc.up.railway.app/status-board">Status Board</a>.</p>')
        ok, err = send_email(MAIL_TO, f'CPAinc Alerts — {today}', body)
        sent_alert = ok
        log.append(f"Alert email: {'sent' if ok else 'FAILED — ' + str(err)}")
    else:
        log.append('No alerts to send today.')

    # ── 3. Sunday DB backup ───────────────────────────────────────────────────
    import sqlite3 as _sq3, tempfile as _tmp
    sent_backup = False
    if datetime.today().weekday() == 6:   # 6 = Sunday
        try:
            with _tmp.NamedTemporaryFile(suffix='.sqlite', delete=False) as tf:
                tmp_path = tf.name
            src = _sq3.connect(DATABASE)
            dst = _sq3.connect(tmp_path)
            src.backup(dst); src.close(); dst.close()
            with open(tmp_path, 'rb') as f:
                db_bytes = f.read()
            os.unlink(tmp_path)
            backup_name = f'CPAinc_backup_{today}.sqlite'
            ok, err = send_email(
                MAIL_TO,
                f'CPAinc Weekly DB Backup — {today}',
                f'<p>Automated weekly database backup attached ({len(db_bytes)//1024} KB).</p>',
                attachments=[(backup_name, db_bytes)]
            )
            sent_backup = ok
            log.append(f"DB backup email: {'sent' if ok else 'FAILED — ' + str(err)}")
        except Exception as e:
            log.append(f'DB backup error: {e}')
    else:
        log.append('DB backup skipped (not Sunday).')

    return jsonify({'ok': True, 'date': today, 'log': log,
                    'alerts_sent': sent_alert, 'backup_sent': sent_backup})


@app.route('/admin/test-email')
def admin_test_email():
    """Send a test email to verify MAIL_* settings are correct."""
    user = get_current_user()
    if not user or user['role'] != 'admin':
        return 'Forbidden', 403
    to = request.args.get('to') or MAIL_TO
    if not to:
        return 'No recipient — set MAIL_TO env var or pass ?to=email@example.com', 400
    ok, err = send_email(to, 'CPAinc — Test Email',
                         '<p>Your CPAinc email alerts are configured correctly!</p>')
    if ok:
        return f'Test email sent to {to} ✓'
    return f'Failed: {err}', 500


if __name__ == '__main__':
    app.run(debug=True, port=5051, use_reloader=False)
