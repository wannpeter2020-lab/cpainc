import os
import io
import json
import sqlite3
import functools
import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, flash, jsonify, send_file, g, session
from werkzeug.security import generate_password_hash, check_password_hash
from datetime import datetime

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
DATABASE = os.path.join(_DATA_DIR, 'CPAinc.sqlite')

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
    """Convert YYYY-MM-DD (or datetime) to MM/DD/YYYY for display."""
    if not value:
        return '—'
    try:
        s = str(value)[:10]
        from datetime import datetime as _dt
        return _dt.strptime(s, '%Y-%m-%d').strftime('%m/%d/%Y')
    except Exception:
        return value

@app.template_filter('fmtdate')
def fmt_date(val):
    if not val:
        return ''
    try:
        s = str(val)[:10]
        from datetime import datetime as _dt
        return _dt.strptime(s, '%Y-%m-%d').strftime('%m/%d/%Y')
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
    bookings  = db.execute(query, params).fetchall()
    statuses  = [r[0] for r in db.execute('SELECT DISTINCT BookingStatus FROM ReportPipeline WHERE BookingStatus IS NOT NULL ORDER BY 1').fetchall()]
    associates = [r[0] for r in db.execute('SELECT DISTINCT BookingAssociate FROM ReportPipeline WHERE BookingAssociate IS NOT NULL ORDER BY 1').fetchall()]

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
    db = get_db()
    booking = db.execute('SELECT * FROM ReportPipeline WHERE BookingId = ?', (booking_id,)).fetchone()
    if not booking:
        flash('Booking not found.', 'error')
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
    added = 0
    for f in files:
        if f and f.filename:
            db.execute(
                'INSERT INTO booking_contract (booking_id, filename, file_data, upload_date) VALUES (?,?,?,date("now"))',
                (booking_id, f.filename, f.read())
            )
            added += 1
    db.commit()
    flash(f'{added} contract file{"s" if added != 1 else ""} uploaded.', 'success')
    return redirect(url_for('booking_detail', booking_id=booking_id))

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
                header_row = 0
                df.columns = [str(c).strip() for c in df.columns]
            else:
                df = pd.read_excel(tmp_path, engine='xlrd' if ext == 'xls' else 'openpyxl', header=None)
                header_row = None
                for i, row in df.iterrows():
                    if any(str(v).strip() in ('Booking Id', 'Booking ID') for v in row.values):
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
                'booking associate: full name':             'Booking Associate',
                'booking associate':                        'Booking Associate',
                'event: event name':                        'Event Name',
                'event: event  name':                       'Event Name',
                'booking name':                             'Booking Name',
                'account name: account name':               'Account Name',
                'account name':                             'Account Name',
                'entity to invoice: account name':          'Customer',
                'entity to invoice':                        'Customer',
                'entity to invoice: billing street':        'Address',
                'entity to invoice: billing city':          'City',
                'entity to invoice: billing state/province': 'State',
                'entity to invoice: billing country':       'Country',
                'status':                                   'Booking Status',
                'type':                                     'Booking Type',
                'subtype':                                  'Booking Type',
                'addendum submitted':                       'Addendum',
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
                'Booking Id', 'Booking Name', 'Booking Associate', 'Share Type', 'Booking Type', 'Booking Status',
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
                    'Booking Id': 'BookingId', 'Booking Name': 'BookingName',
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
                    # Update Commission Percent if previously missing
                    comm_idx = fields.index('Commission Percent')
                    if values[comm_idx] is not None:
                        db.execute(
                            'UPDATE ReportPipeline SET CommissionPercent = ? WHERE BookingId = ? AND (CommissionPercent IS NULL OR CommissionPercent = "")',
                            (values[comm_idx], bid)
                        )
                        if db.execute('SELECT changes()').fetchone()[0]:
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
            flash(f'Import complete: {added} added, {updated} commission % updated, {skipped} already existed (skipped).', 'success')
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
        with pdfplumber.open(_io.BytesIO(raw_bytes)) as pdf:
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

import sys as _sys_cs, os as _os_cs
_sys_cs.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
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

    conn.close()

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
        return redirect(url_for('settings'))

    current       = get_commission_split()
    tolerance     = get_payment_tolerance()
    kristin_split = get_kristin_split()
    kristin_cut   = get_kristin_cut()
    overrides = db.execute('SELECT id, account_name, split_rate, countries FROM AccountSplits ORDER BY account_name').fetchall()
    accounts  = [r[0] for r in db.execute('SELECT DISTINCT AccountName FROM ReportPipeline WHERE AccountName IS NOT NULL ORDER BY AccountName').fetchall()]
    return render_template('settings.html', commission_split=current, tolerance=tolerance,
                           kristin_split=kristin_split, kristin_cut=kristin_cut,
                           overrides=overrides, accounts=accounts)

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
                if len(nums) >= 3:
                    result['avg_rate'] = nums[2]
                break

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
    # Add BookingName to ReportPipeline if not present
    try:
        db.execute('ALTER TABLE ReportPipeline ADD COLUMN BookingName TEXT')
        db.commit()
    except Exception:
        pass


with app.app_context():
    ensure_pickup_tables()


# ── Auth tables & seeding ─────────────────────────────────────────────────────

ALL_PERMISSIONS = [
    ('dashboard',                  'Dashboard'),
    ('admin_panel',                'Admin Panel'),
    ('import_bookings',            'Import → Bookings'),
    ('import_payments',            'Import → Payments'),
    ('import_hhr',                 'Import → HHR'),
    ('import_cancelled',           'Import → Cancelled Meetings'),
    ('reports_commission_kristin', 'Reports → Kristin Commission'),
    ('reports_commission_team',    'Reports → Team Commission'),
    ('reports_payments',           'Reports → Payment Report'),
    ('reports_customer_summary',   'Reports → Customer Summary'),
    ('bookings_view',              'View Bookings'),
    ('bookings_edit',              'Add / Edit Bookings'),
    ('pickups_payments',           'Add Pickups / Payments'),
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
    ''')
    db.commit()
    _seed_users(db)


def _hash_password(password):
    return generate_password_hash(password, method='pbkdf2:sha256')


def _seed_users(db):
    for name, email, username, password, role in _SEED_USERS:
        existing = db.execute('SELECT id FROM Users WHERE email = ?', (email,)).fetchone()
        if not existing:
            ph = _hash_password(password)
            db.execute(
                'INSERT INTO Users (name, email, username, password_hash, role) VALUES (?,?,?,?,?)',
                (name, email, username, ph, role)
            )
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


with app.app_context():
    ensure_auth_tables()


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
    """Returns a list of account names the user may see, or None for admins (no filter)."""
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
    user = get_db().execute('SELECT id, active FROM Users WHERE id = ?', (uid,)).fetchone()
    if not user or not user['active']:
        session.clear()
        return redirect(url_for('login'))


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
            session['user_id'] = user['id']
            next_page = request.form.get('next') or request.args.get('next') or url_for('index')
            return redirect(next_page)
        error = 'Invalid username or password.'
    return render_template('login.html', error=error, next=request.args.get('next', ''))


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
        "AND LOWER(BookingAssociate) IN ('kristin house', 'morgan basham') "
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

    db    = get_db()
    today = datetime.today().strftime('%Y-%m-%d')

    # All future non-cancelled bookings without a pickup_config entry
    missing = db.execute('''
        SELECT r.BookingId, r.BookingName, r.EventName, r.AccountName,
               r.Customer, r.StartDate, r.EndDate, r.PeakRooms, r.RoomRate,
               r.BookingAssociate, r.BookingStatus
        FROM ReportPipeline r
        WHERE (r.BookingStatus IS NULL OR r.BookingStatus NOT LIKE '%Cancel%')
          AND r.EndDate >= ?
          AND NOT EXISTS (
              SELECT 1 FROM pickup_config p WHERE p.booking_id = CAST(r.BookingId AS TEXT)
          )
        ORDER BY r.StartDate
    ''', (today,)).fetchall()

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


@app.route('/pickup')
def pickup_dashboard():
    from datetime import date, timedelta
    db = get_db()
    today = date.today()
    today_str = today.isoformat()
    future_cutoff = today + timedelta(days=120)
    configs = db.execute("SELECT * FROM pickup_config WHERE status='active' ORDER BY cutoff_date").fetchall()
    archived_configs = db.execute("SELECT * FROM pickup_config WHERE status='archived' ORDER BY cutoff_date").fetchall()
    past_rows, current_rows, future_rows, archived_rows = [], [], [], []

    for c in configs:
        last = db.execute(
            "SELECT * FROM pickup_weekly WHERE config_id=? ORDER BY report_date DESC LIMIT 1", (c['id'],)
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
        last_date      = last['report_date'] if last else None
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
        event_start = all_dates[0] if all_dates else None
        event_end   = all_dates[-1] if all_dates else None
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
        if all_dates:
            try:
                start_date = date.fromisoformat(all_dates[0])
            except Exception:
                pass
        if has_final_history or force_past:
            past_rows.append(row)
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
            'event_start': all_dates[0] if all_dates else None,
            'event_end':   all_dates[-1] if all_dates else None,
        })

    return render_template('pickup_dashboard.html',
                           past_rows=past_rows, current_rows=current_rows,
                           future_rows=future_rows, archived_rows=archived_rows,
                           today=today_str)


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
        # Dynamic hotel contacts (hc_name_N / hc_email_N)
        hotel_contacts = []
        for i in range(20):
            e_val = f.get(f'hc_email_{i}', '').strip()
            n_val = f.get(f'hc_name_{i}', '').strip()
            if e_val or n_val:
                hotel_contacts.append({'name': n_val, 'email': e_val})
        primary_hc_name  = hotel_contacts[0]['name']  if hotel_contacts else None
        primary_hc_email = hotel_contacts[0]['email'] if hotel_contacts else None
        # Dynamic client contacts (cc_name_N / cc_email_N)
        cc_emails = []
        for i in range(20):
            e_val = f.get(f'cc_email_{i}', '').strip()
            n_val = f.get(f'cc_name_{i}', '').strip()
            if e_val or n_val:
                cc_emails.append({'name': n_val, 'email': e_val})
        primary_gc_name  = cc_emails[0]['name']  if cc_emails else None
        primary_gc_email = cc_emails[0]['email'] if cc_emails else None
        db = get_db()
        db.execute('''
            INSERT INTO pickup_config
            (booking_id, tab_name, organization, event_name, hotel, hotel_contact,
             hotel_contact_email, hotel_contacts, group_contact, group_contact_email,
             cutoff_date, attrition_pct, contracted_block, contracted_rate, shoulder_pre,
             shoulder_post, hotel_booking_link, notes, ota_url, cc_emails)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        ''', (
            f.get('booking_id'), f.get('tab_name'), f['organization'],
            f.get('event_name'), f.get('hotel'), primary_hc_name,
            primary_hc_email, json.dumps(hotel_contacts), primary_gc_name,
            primary_gc_email, f.get('cutoff_date'),
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
    weekly = db.execute("SELECT * FROM pickup_weekly WHERE config_id=? ORDER BY report_date DESC", (cid,)).fetchall()
    rooming = db.execute(
        "SELECT id, upload_date, filename, total_guests, reconciliation_status, discrepancy_notes FROM pickup_rooming_list WHERE config_id=? ORDER BY upload_date DESC", (cid,)
    ).fetchall()
    contact_log = db.execute(
        "SELECT * FROM pickup_contact_log WHERE config_id=? ORDER BY contact_date DESC, id DESC", (cid,)
    ).fetchall()
    block = json.loads(config['contracted_block'] or '{}')
    all_dates = sorted(block.keys())
    contracted_total = sum(v for v in block.values() if v)
    day_map = {}
    for d in all_dates:
        try:
            day_map[d] = _dt.strptime(d, '%Y-%m-%d').strftime('%a')
        except Exception:
            day_map[d] = ''
    attrition_pct   = config['attrition_pct'] or 0
    attrition_rooms = round(contracted_total * attrition_pct, 1)
    weekly_display = []
    prev_total = None
    for w in reversed(weekly):
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
    has_final_history = any(w['label'] and 'final' in w['label'].lower() for w in weekly)
    if all_dates:
        try:
            if _dt.strptime(all_dates[0], '%Y-%m-%d') < _dt.today():
                past_cutoff = True
        except Exception:
            pass
    return render_template('pickup_event.html',
                           config=config, weekly=weekly_display, rooming=rooming,
                           contact_log=contact_log, block=block, all_dates=all_dates,
                           day_map=day_map, contracted_total=contracted_total,
                           attrition_pct=attrition_pct, attrition_rooms=attrition_rooms,
                           past_cutoff=past_cutoff, has_final_history=has_final_history,
                           today=_date.today().isoformat())


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
                group_contact       = request.form.get('group_contact', '').strip() or None
                group_contact_email = request.form.get('group_contact_email', '').strip() or None

                contract_data     = request.form.get('_contract_data_b64', '')
                contract_filename = request.form.get('_contract_filename', '')
                import base64
                file_blob = base64.b64decode(contract_data) if contract_data else None

                db.execute('''
                    UPDATE pickup_config
                    SET contracted_block    = ?,
                        contracted_rate     = ?,
                        cutoff_date         = ?,
                        attrition_pct       = ?,
                        block_is_estimated  = 0,
                        contract_filename   = ?,
                        contract_data       = ?,
                        hotel_contact       = COALESCE(?, hotel_contact),
                        hotel_contact_email = COALESCE(?, hotel_contact_email),
                        group_contact       = COALESCE(?, group_contact),
                        group_contact_email = COALESCE(?, group_contact_email)
                    WHERE id = ?
                ''', (json.dumps(block), rate, cutoff, atr,
                      contract_filename or None, file_blob,
                      hotel_contact, hotel_contact_email,
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
            return redirect(url_for('pickup_upload_contract', cid=cid))

        import base64
        file_b64 = base64.b64encode(file_bytes).decode('utf-8')
        return render_template('pickup_contract_review.html',
                               config=config, extracted=extracted,
                               filename=f.filename, file_b64=file_b64)

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
    buf = _io.BytesIO(row['contract_data'])
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

    buf = _io.BytesIO()
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


@app.route('/pickup/<int:cid>/edit', methods=['GET', 'POST'])
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
        # Dynamic hotel contacts
        hotel_contacts = []
        for i in range(20):
            e_val = f.get(f'hc_email_{i}', '').strip()
            n_val = f.get(f'hc_name_{i}', '').strip()
            if e_val or n_val:
                hotel_contacts.append({'name': n_val, 'email': e_val})
        primary_hc_name  = hotel_contacts[0]['name']  if hotel_contacts else None
        primary_hc_email = hotel_contacts[0]['email'] if hotel_contacts else None
        # Dynamic client contacts
        cc_emails = []
        for i in range(20):
            e_val = f.get(f'cc_email_{i}', '').strip()
            n_val = f.get(f'cc_name_{i}', '').strip()
            if e_val or n_val:
                cc_emails.append({'name': n_val, 'email': e_val})
        primary_gc_name  = cc_emails[0]['name']  if cc_emails else None
        primary_gc_email = cc_emails[0]['email'] if cc_emails else None
        db.execute('''
            UPDATE pickup_config SET
            booking_id=?, tab_name=?, organization=?, event_name=?, hotel=?,
            hotel_contact=?, hotel_contact_email=?, hotel_contacts=?,
            group_contact=?, group_contact_email=?, cutoff_date=?, attrition_pct=?,
            contracted_block=?, contracted_rate=?, shoulder_pre=?,
            shoulder_post=?, hotel_booking_link=?, notes=?, ota_url=?, cc_emails=?
            WHERE id=?
        ''', (
            f.get('booking_id'), f.get('tab_name'), f['organization'],
            f.get('event_name'), f.get('hotel'), primary_hc_name,
            primary_hc_email, json.dumps(hotel_contacts), primary_gc_name,
            primary_gc_email, f.get('cutoff_date'),
            attrition, json.dumps(contracted_block),
            float(f['contracted_rate']) if f.get('contracted_rate') else None,
            int(f.get('shoulder_pre', 3)), int(f.get('shoulder_post', 3)),
            f.get('hotel_booking_link'), f.get('notes'), ota_url, json.dumps(cc_emails), cid
        ))
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
        for d in block:
            val = f.get(f'night_{d}', '')
            if val != '':
                pickup_by_night[d] = int(val)
        total_rooms = sum(pickup_by_night.values())
        contracted_total = sum(block.values())
        change_from_last = (total_rooms - (last['total_rooms'] or 0)) if last else None
        pct_of_block     = round(total_rooms / contracted_total * 100, 1) if contracted_total else None
        attrition_floor  = contracted_total * (config['attrition_pct'] or 0)
        pct_of_attrition = round(total_rooms / attrition_floor * 100, 1) if attrition_floor else None
        ota_rate = float(f['ota_rate']) if f.get('ota_rate') else None
        db.execute('''
            INSERT INTO pickup_weekly
            (config_id, report_date, pickup_by_night, total_rooms, change_from_last,
             pct_of_block, pct_of_attrition, ota_rate, label, notes)
            VALUES (?,?,?,?,?,?,?,?,?,?)
        ''', (cid, f['report_date'], json.dumps(pickup_by_night), total_rooms,
              change_from_last, pct_of_block, pct_of_attrition, ota_rate,
              f.get('label'), f.get('notes')))
        db.commit()
        flash('Weekly entry saved.', 'success')
        return redirect(url_for('pickup_event', cid=cid))
    block = json.loads(config['contracted_block'] or '{}')
    last_pickup = json.loads(last['pickup_by_night']) if last else {}
    today_str = datetime.today().strftime('%Y-%m-%d')
    return render_template('pickup_weekly_form.html', config=config, block=block,
                           last=last, last_pickup=last_pickup, today=today_str,
                           is_edit=False, entry=None, entry_pickup={}, wid=None)


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
        for d in block:
            val = f.get(f'night_{d}', '')
            if val != '':
                pickup_by_night[d] = int(val)
        total_rooms = sum(pickup_by_night.values())
        contracted_total = sum(block.values())
        pct_of_block     = round(total_rooms / contracted_total * 100, 1) if contracted_total else None
        attrition_floor  = contracted_total * (config['attrition_pct'] or 0)
        pct_of_attrition = round(total_rooms / attrition_floor * 100, 1) if attrition_floor else None
        ota_rate = float(f['ota_rate']) if f.get('ota_rate') else None
        db.execute('''
            UPDATE pickup_weekly
            SET report_date=?, pickup_by_night=?, total_rooms=?,
                pct_of_block=?, pct_of_attrition=?, ota_rate=?, label=?, notes=?
            WHERE id=? AND config_id=?
        ''', (f['report_date'], json.dumps(pickup_by_night), total_rooms,
              pct_of_block, pct_of_attrition, ota_rate,
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
    prev_entry = db.execute(
        "SELECT * FROM pickup_weekly WHERE config_id=? AND report_date < ? ORDER BY report_date DESC LIMIT 1",
        (cid, entry['report_date'])
    ).fetchone()
    last_pickup = json.loads(prev_entry['pickup_by_night']) if prev_entry and prev_entry['pickup_by_night'] else {}
    return render_template('pickup_weekly_form.html',
                           config=config, block=block,
                           last=prev_entry, last_pickup=last_pickup,
                           entry_pickup=entry_pickup, entry=entry,
                           today=entry['report_date'], is_edit=True, wid=wid)


@app.route('/pickup/<int:cid>/weekly/<int:wid>/delete', methods=['POST'])
def pickup_weekly_delete(cid, wid):
    db = get_db()
    db.execute("DELETE FROM pickup_weekly WHERE id=? AND config_id=?", (wid, cid))
    db.commit()
    flash('Weekly entry deleted.', 'success')
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


@app.route('/pickup/<int:cid>/email/hotel')
def pickup_email_hotel(cid):
    from pickup_utils import build_hotel_email
    db = get_db()
    config = db.execute("SELECT * FROM pickup_config WHERE id=?", (cid,)).fetchone()
    if not config:
        flash('Event not found.', 'error')
        return redirect(url_for('pickup_dashboard'))
    email = build_hotel_email(config)
    return render_template('pickup_email_preview.html', config=config, email=email, email_type='hotel')


@app.route('/pickup/<int:cid>/email/client')
def pickup_email_client(cid):
    from pickup_utils import build_client_email
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
    email = build_client_email(config, last, rl_status, weekly_list=weekly_dicts)
    return render_template('pickup_email_preview.html', config=config, email=email,
                           email_type='client', rooming_list=rl)


@app.route('/pickup/<int:cid>/email/client/launch-outlook')
def pickup_email_client_launch_outlook(cid):
    import subprocess, tempfile
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
    email = build_client_email(config, last, rl_status, weekly_list=weekly_dicts)
    html_body     = email.get('html_body', '')
    subject       = (email.get('subject') or '').strip()
    to_addr       = (email.get('to') or '').strip()
    cc_recipients = _build_cc_recipients(config)

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

    to_line = (f'make new to recipient at theMsg with properties {{email address:{{name:"", address:"{esc(to_addr)}"}}}}' if to_addr else '')
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
        subprocess.run(['osascript', '-l', 'JavaScript', clip_path], check=True)
        subprocess.Popen(['osascript', outlook_path])
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
    db = get_db()
    show_archived = request.args.get('archived') == '1'
    rfps = db.execute(
        'SELECT r.*, (SELECT COUNT(*) FROM rfp_hotel h WHERE h.rfp_id=r.id) AS hotel_count '
        'FROM rfp r WHERE r.archived=? ORDER BY r.created_at DESC',
        (1 if show_archived else 0,)
    ).fetchall()
    return render_template('rfp_dashboard.html', rfps=rfps, statuses=RFP_STATUS_MAP,
                           show_archived=show_archived, all_statuses=RFP_STATUSES)


@app.route('/rfp/new', methods=['GET', 'POST'])
def rfp_new():
    if request.method == 'POST':
        db = get_db()
        db.execute('''
            INSERT INTO rfp (client_org, event_name, rfp_name, start_date, end_date,
                alt_start_date, alt_end_date, peak_rooms, total_room_nights,
                total_attendees, f_and_b_budget, response_due_date, decision_due_date,
                status, notes)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        ''', (
            request.form.get('client_org', '').strip(),
            request.form.get('event_name', '').strip() or None,
            request.form.get('rfp_name', '').strip() or None,
            request.form.get('start_date') or None,
            request.form.get('end_date') or None,
            request.form.get('alt_start_date') or None,
            request.form.get('alt_end_date') or None,
            request.form.get('peak_rooms') or None,
            request.form.get('total_room_nights') or None,
            request.form.get('total_attendees') or None,
            request.form.get('f_and_b_budget') or None,
            request.form.get('response_due_date') or None,
            request.form.get('decision_due_date') or None,
            request.form.get('status', 'sourcing'),
            request.form.get('notes', '').strip() or None,
        ))
        db.commit()
        flash('RFP created.', 'success')
        new_id = db.execute('SELECT last_insert_rowid()').fetchone()[0]
        return redirect(url_for('rfp_detail', rid=new_id))
    return render_template('rfp_form.html', rfp=None, all_statuses=RFP_STATUSES)


@app.route('/rfp/<int:rid>')
def rfp_detail(rid):
    db = get_db()
    rfp = db.execute('SELECT * FROM rfp WHERE id=?', (rid,)).fetchone()
    if not rfp:
        flash('RFP not found.', 'error')
        return redirect(url_for('rfp_dashboard'))
    hotels = db.execute('SELECT * FROM rfp_hotel WHERE rfp_id=? ORDER BY id', (rid,)).fetchall()
    notes = db.execute('SELECT * FROM rfp_note WHERE rfp_id=? ORDER BY note_date DESC', (rid,)).fetchall()
    return render_template('rfp_detail.html', rfp=rfp, hotels=hotels, notes=notes,
                           statuses=RFP_STATUS_MAP, all_statuses=RFP_STATUSES)


@app.route('/rfp/<int:rid>/edit', methods=['GET', 'POST'])
def rfp_edit(rid):
    db = get_db()
    rfp = db.execute('SELECT * FROM rfp WHERE id=?', (rid,)).fetchone()
    if not rfp:
        flash('RFP not found.', 'error')
        return redirect(url_for('rfp_dashboard'))
    if request.method == 'POST':
        db.execute('''
            UPDATE rfp SET client_org=?, event_name=?, rfp_name=?, start_date=?,
                end_date=?, alt_start_date=?, alt_end_date=?, peak_rooms=?,
                total_room_nights=?, total_attendees=?, f_and_b_budget=?,
                response_due_date=?, decision_due_date=?, status=?, notes=?,
                updated_at=datetime('now')
            WHERE id=?
        ''', (
            request.form.get('client_org', '').strip(),
            request.form.get('event_name', '').strip() or None,
            request.form.get('rfp_name', '').strip() or None,
            request.form.get('start_date') or None,
            request.form.get('end_date') or None,
            request.form.get('alt_start_date') or None,
            request.form.get('alt_end_date') or None,
            request.form.get('peak_rooms') or None,
            request.form.get('total_room_nights') or None,
            request.form.get('total_attendees') or None,
            request.form.get('f_and_b_budget') or None,
            request.form.get('response_due_date') or None,
            request.form.get('decision_due_date') or None,
            request.form.get('status', 'sourcing'),
            request.form.get('notes', '').strip() or None,
            rid,
        ))
        db.commit()
        flash('RFP updated.', 'success')
        return redirect(url_for('rfp_detail', rid=rid))
    return render_template('rfp_form.html', rfp=rfp, all_statuses=RFP_STATUSES)


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
    db.execute("UPDATE rfp_hotel SET status='pending' WHERE rfp_id=?", (rid,))
    db.execute("UPDATE rfp_hotel SET status='selected' WHERE id=? AND rfp_id=?", (hid, rid))
    db.execute("UPDATE rfp SET status='hotel_selected', updated_at=datetime('now') WHERE id=?", (rid,))
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


if __name__ == '__main__':
    app.run(debug=True, port=5051, use_reloader=False)
