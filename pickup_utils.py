"""
pickup_utils.py — Utility functions for the Pickup Tracking module.

Functions:
  import_pickup_tab(ws, tab_name)         Parse one Excel tab → config + weekly entries
  parse_rooming_list_pdf(file_bytes)      Extract guests from PDF → nights_by_date
  reconcile(pickup_by_night, rl_nights)   Compare pickup vs rooming list
  format_export_row(config, weekly)       Tab-separated row for Excel paste-back
  hotel_request_mailto(config)            Pre-filled mailto: for hotel pickup request
  client_summary_mailto(config, weekly, rl_status)  mailto: for client summary
"""

import re
import json
import io
from datetime import datetime, timedelta
from urllib.parse import quote


# ── Spreadsheet importer ──────────────────────────────────────────────────────

def _shared_helpers():
    """Return (sv, to_iso) helpers — defined once, reused by both parsers."""
    def sv(v):
        return str(v).strip() if v is not None else ''

    def to_iso(v):
        if v is None:
            return None
        if hasattr(v, 'strftime'):
            return v.strftime('%Y-%m-%d')
        s = str(v).strip()
        if re.match(r'^\d{4}-\d{2}-\d{2}', s):
            return s[:10]
        if re.match(r'^\d{2}/\d{2}/\d{2}$', s):
            m, d, y = s.split('/')
            yr = 2000 + int(y) if int(y) < 50 else 1900 + int(y)
            return f'{yr}-{m}-{d}'
        return None

    return sv, to_iso


def _parse_section(rows, tab_name):
    """
    Parse a list of rows that represents one hotel's pickup table.
    rows[0] must be the ORGANIZATION: row (col A = 'ORGANIZATION:' or similar).

    Layout (0-indexed within the slice):
      0  ORGANIZATION:      | org_name
      1  HOTEL & LOCATION:  | hotel_name
      2  NAME & DATE …      | event_name
      3  GROUP CONTACT      | gc_name  [col C may hold gc email]
      4  HOTEL CONTACT:     | hc_name_or_email
      5  CONTACT'S NUMBER:  | phone_or_email
      6  Booking            | booking_id
      7  Cut-off Date:      | cutoff  …  Attrition: | pct
      8  Dates:             | d1 | d2 | …
      9  Day:               | Mon | Tue | …
     10  Block:             | b1 | b2 | …
     11+ weekly rows
    """
    sv, to_iso = _shared_helpers()

    def cell(r, c):
        try:
            return rows[r][c]
        except (IndexError, TypeError):
            return None

    def _int(v):
        try:
            return int(v) if v not in (None, '', '#DIV/0!', 'NA', 'N/A') else None
        except Exception:
            return None

    def _flt(v):
        try:
            return float(v) if v not in (None, '', '#DIV/0!', 'NA', 'N/A') else None
        except Exception:
            return None

    organization = sv(cell(0, 1))
    hotel        = sv(cell(1, 1))
    event_name   = sv(cell(2, 1))

    # Group contact: col B; email sometimes in col C
    group_contact = sv(cell(3, 1))
    gc_col_c = sv(cell(3, 2))
    group_contact_email = gc_col_c if '@' in gc_col_c else ''

    # Hotel contact: name in col B; email may be directly in col B, or in row 5
    hc_raw = sv(cell(4, 1))
    hotel_contact_email = ''
    em = re.search(r'<([^>]+@[^>]+)>', hc_raw)
    if em:
        hotel_contact_email = em.group(1)
        hotel_contact = hc_raw[:hc_raw.find('<')].strip()
    elif '@' in hc_raw:
        hotel_contact_email = hc_raw
        hotel_contact = ''
    else:
        hotel_contact = hc_raw
    # Fallback: email in row 5 (CONTACT'S NUMBER)
    if not hotel_contact_email:
        row5 = sv(cell(5, 1))
        if '@' in row5:
            hotel_contact_email = row5

    # Booking ID
    bid_raw = cell(6, 1)
    try:
        booking_id = str(int(float(str(bid_raw).strip()))) if bid_raw else None
    except Exception:
        booking_id = sv(bid_raw) or None

    # Cut-off date and attrition (row 7)
    cutoff_date = to_iso(cell(7, 1))
    attrition_pct = None
    if len(rows) > 7 and rows[7]:
        for v in rows[7]:
            if isinstance(v, (int, float)) and 0 < v <= 1:
                attrition_pct = float(v)
                break

    # Event dates (row 8, cols B+)
    event_dates = []
    if len(rows) > 8 and rows[8]:
        for v in (rows[8][1:] if len(rows[8]) > 1 else []):
            d = to_iso(v)
            if d:
                event_dates.append(d)
            elif v is None and event_dates:
                break
    n_nights = len(event_dates)

    # Contracted block (row 10, cols B+)
    contracted_block = {}
    if len(rows) > 10 and n_nights:
        for i, date in enumerate(event_dates):
            v = cell(10, i + 1)
            try:
                contracted_block[date] = int(v) if v not in (None, '', 'N/A') else 0
            except Exception:
                contracted_block[date] = 0

    # Weekly entries (row 11+)
    _SKIP = {'remaining', 'pick-up update', 'balance',
             'revised block', 'day', 'dates', 'block', 'organization',
             'hotel & location', 'name & date of event', 'group contact',
             'hotel contact', "contact's number", 'booking', 'cut-off date',
             'attrition'}
    LABEL_PREFIXES = ('final history', 'final pickup', 'pending history',
                      'pending pickup', 'revised')
    weekly_entries = []
    rows_list = list(rows)   # ensure indexable for .index()

    for ri, row in enumerate(rows_list[11:], start=11):
        if not row:
            continue
        col_a = row[0]
        if col_a is None:
            continue

        label = None
        report_date = None

        if isinstance(col_a, str):
            stripped = col_a.strip()
            stripped_lo = stripped.lower()
            if not stripped or stripped_lo in _SKIP:
                continue

            # Check for label prefix rows (Final History, Pending History, etc.)
            matched_prefix = None
            for pfx in LABEL_PREFIXES:
                if stripped_lo.startswith(pfx):
                    matched_prefix = pfx
                    break

            if matched_prefix is not None:
                label = stripped  # preserve original casing
                # Try to extract a date from remainder of text or col B
                remainder = stripped[len(matched_prefix):].strip().strip(':-')
                report_date = to_iso(remainder) or to_iso(cell(ri, 1))
                # For label rows without a parseable date, use last event date
                if not report_date and event_dates:
                    report_date = event_dates[-1]
                if not report_date:
                    continue
            else:
                report_date = to_iso(stripped)
                if not report_date:
                    continue
        elif hasattr(col_a, 'strftime'):
            report_date = col_a.strftime('%Y-%m-%d')
        else:
            report_date = to_iso(col_a)
            if not report_date:
                continue

        if not report_date:
            continue

        pickup_by_night = {}
        if n_nights:
            for i, date in enumerate(event_dates):
                v = row[i + 1] if (i + 1) < len(row) else None
                try:
                    pickup_by_night[date] = int(v) if v not in (None, '') else None
                except Exception:
                    pickup_by_night[date] = None

        col_total   = n_nights + 1
        col_change  = n_nights + 2
        col_pct_blk = n_nights + 3
        col_pct_att = n_nights + 4
        col_ota     = n_nights + 5

        total_rooms          = _int(row[col_total]   if col_total   < len(row) else None)
        change_from_last     = _int(row[col_change]  if col_change  < len(row) else None)
        pct_of_block_raw     = _flt(row[col_pct_blk] if col_pct_blk < len(row) else None)
        pct_of_attrition_raw = _flt(row[col_pct_att] if col_pct_att < len(row) else None)
        pct_of_block     = round(pct_of_block_raw * 100, 1) if (pct_of_block_raw is not None and pct_of_block_raw <= 1.5) else pct_of_block_raw
        pct_of_attrition = round(pct_of_attrition_raw * 100, 1) if (pct_of_attrition_raw is not None and pct_of_attrition_raw <= 1.5) else pct_of_attrition_raw
        ota_rate         = _flt(row[col_ota]     if col_ota     < len(row) else None)

        if not pickup_by_night and total_rooms is None:
            continue

        weekly_entries.append({
            'report_date':      report_date,
            'pickup_by_night':  json.dumps(pickup_by_night),
            'total_rooms':      total_rooms,
            'change_from_last': change_from_last,
            'pct_of_block':     pct_of_block,
            'pct_of_attrition': pct_of_attrition,
            'ota_rate':         ota_rate,
            'label':            label,
        })

    return {
        'config': {
            'tab_name':             tab_name,
            'organization':         organization,
            'event_name':           event_name,
            'hotel':                hotel,
            'hotel_contact':        hotel_contact,
            'hotel_contact_email':  hotel_contact_email,
            'group_contact':        group_contact,
            'group_contact_email':  group_contact_email,
            'booking_id':           booking_id,
            'cutoff_date':          cutoff_date,
            'attrition_pct':        attrition_pct,
            'contracted_block':     contracted_block,
        },
        'weekly_entries': weekly_entries,
    }


def import_pickup_tab_multi(ws, tab_name):
    """
    Parse one openpyxl worksheet that may contain multiple hotel tables.

    Each table starts with a row where col A = 'ORGANIZATION:'.
    Returns a list of dicts: [{config, weekly_entries}, ...] — one per table.
    Skips summary/rollup tables that have no contracted_block and no weekly entries.
    """
    sv, to_iso = _shared_helpers()
    all_rows = list(ws.iter_rows(values_only=True))

    # Find every row where col A = 'ORGANIZATION:' (case-insensitive, strip colon)
    org_indices = [
        i for i, row in enumerate(all_rows)
        if row and sv(row[0]).upper().rstrip(':').strip() == 'ORGANIZATION'
    ]

    if not org_indices:
        return []

    # Split at each ORGANIZATION: boundary
    sections = []
    for n, start in enumerate(org_indices):
        end = org_indices[n + 1] if n + 1 < len(org_indices) else len(all_rows)
        sections.append(all_rows[start:end])

    results = []
    for n, section_rows in enumerate(sections):
        result = _parse_section(section_rows, tab_name)
        if result is None:
            continue
        cfg = result['config']
        # Skip pure rollup sections: no booking_id AND no weekly entries AND
        # hotel name looks like a summary ("All Hotels", "Total", etc.)
        hotel_lo = (cfg.get('hotel') or '').lower()
        is_rollup = (not cfg.get('booking_id')
                     and not result['weekly_entries']
                     and (not cfg.get('contracted_block') or
                          hotel_lo in ('all hotels', 'all', 'total', 'totals', '')))
        if is_rollup:
            continue
        if not cfg.get('organization'):
            continue
        # Tag with hotel suffix when multiple tables exist
        if len(sections) > 1:
            hotel_tag = cfg.get('hotel') or f'Table {n+1}'
            cfg['tab_name'] = f"{tab_name} [{hotel_tag}]"
        results.append(result)

    return results


def import_pickup_tab(ws, tab_name):
    """
    Parse one openpyxl worksheet tab from the pickup spreadsheet.

    Returns a dict:
      {
        'config': {organization, event_name, hotel, hotel_contact,
                   hotel_contact_email, group_contact, booking_id,
                   cutoff_date, attrition_pct, contracted_block (JSON str)},
        'weekly_entries': [{report_date, pickup_by_night (JSON str),
                            total_rooms, change_from_last, pct_of_block,
                            pct_of_attrition, ota_rate, label}]
      }
    """
    rows = list(ws.iter_rows(values_only=True))

    def cell(r, c):
        """0-indexed row, 0-indexed col."""
        try:
            return rows[r][c]
        except IndexError:
            return None

    def sv(v):
        """Safe string — strip whitespace, return '' for None."""
        return str(v).strip() if v is not None else ''

    def to_iso(v):
        """Convert datetime / 'YYYY-MM-DD HH:MM:SS' string / None to ISO date string."""
        if v is None:
            return None
        if hasattr(v, 'strftime'):
            return v.strftime('%Y-%m-%d')
        s = str(v).strip()
        # Accept 'YYYY-MM-DD ...' or 'MM/DD/YY'
        if re.match(r'^\d{4}-\d{2}-\d{2}', s):
            return s[:10]
        if re.match(r'^\d{2}/\d{2}/\d{2}$', s):
            m, d, y = s.split('/')
            yr = 2000 + int(y) if int(y) < 50 else 1900 + int(y)
            return f'{yr}-{m}-{d}'
        return None

    # ── Header fields (rows 0-11, 0-indexed) ─────────────────────────────────
    organization  = sv(cell(1, 1))
    hotel         = sv(cell(2, 1))
    event_name    = sv(cell(3, 1))

    # Group contact may span cols B + C
    gc = ' '.join(sv(cell(4, c)) for c in range(1, 3)).strip()
    group_contact = gc

    # Hotel contact: may contain email in angle brackets
    hc_raw = sv(cell(5, 1))
    hotel_contact_email = ''
    em = re.search(r'<([^>]+@[^>]+)>', hc_raw)
    if em:
        hotel_contact_email = em.group(1)
        hotel_contact = hc_raw[:hc_raw.find('<')].strip()
    else:
        hotel_contact = hc_raw

    # Booking ID
    bid_raw = cell(7, 1)
    try:
        booking_id = str(int(float(str(bid_raw).strip()))) if bid_raw else None
    except Exception:
        booking_id = sv(bid_raw) or None

    # Cut-off date (row 8 col B)
    cutoff_date = to_iso(cell(8, 1))

    # Attrition % — find the first float between 0 and 1 in row 8
    attrition_pct = None
    if len(rows) > 8:
        for v in rows[8]:
            if isinstance(v, (int, float)) and 0 < v < 1:
                attrition_pct = float(v)
                break

    # Event dates (row 9, cols B onwards until None)
    event_dates = []
    if len(rows) > 9:
        for v in rows[9][1:]:
            d = to_iso(v)
            if d:
                event_dates.append(d)
            elif v is None and event_dates:
                break   # stop at first gap after dates started

    n_nights = len(event_dates)

    # Contracted block (row 11, cols B .. B+n_nights-1)
    contracted_block = {}
    if len(rows) > 11 and n_nights:
        for i, date in enumerate(event_dates):
            v = cell(11, i + 1)
            try:
                contracted_block[date] = int(v) if v not in (None, '', 'N/A') else 0
            except Exception:
                contracted_block[date] = 0

    # ── Weekly pickup rows (row 12 onwards) ──────────────────────────────────
    _SKIP_LABELS = {'Remaining', 'Final History', 'PICK-UP UPDATE',
                    'remaining', 'final history'}
    weekly_entries = []

    for row in rows[12:]:
        if not row:
            continue
        col_a = row[0]
        if col_a is None:
            continue

        # Determine label and report_date
        label = None
        report_date = None

        if isinstance(col_a, str):
            stripped = col_a.strip()
            if stripped in _SKIP_LABELS or not stripped:
                continue
            if stripped.lower() == 'final history':
                label = 'Final History'
                # date may be in col B for 'Final History'
                report_date = to_iso(cell(rows.index(row), 1)) if hasattr(rows, 'index') else None
                if not report_date:
                    continue
            else:
                report_date = to_iso(stripped)
                if not report_date:
                    continue
        elif hasattr(col_a, 'strftime'):
            report_date = col_a.strftime('%Y-%m-%d')
        else:
            report_date = to_iso(col_a)
            if not report_date:
                continue

        if not report_date:
            continue

        # Pickup per night (cols 1 .. n_nights)
        pickup_by_night = {}
        if n_nights:
            for i, date in enumerate(event_dates):
                v = row[i + 1] if (i + 1) < len(row) else None
                try:
                    pickup_by_night[date] = int(v) if v not in (None, '') else None
                except Exception:
                    pickup_by_night[date] = None

        def _int(v):
            try:
                return int(v) if v not in (None, '', '#DIV/0!', 'NA', 'N/A') else None
            except Exception:
                return None

        def _flt(v):
            try:
                return float(v) if v not in (None, '', '#DIV/0!', 'NA', 'N/A') else None
            except Exception:
                return None

        col_total    = n_nights + 1
        col_change   = n_nights + 2
        col_pct_blk  = n_nights + 3
        col_pct_att  = n_nights + 4
        col_ota      = n_nights + 5

        total_rooms      = _int(row[col_total]   if col_total   < len(row) else None)
        change_from_last = _int(row[col_change]  if col_change  < len(row) else None)
        pct_of_block_raw     = _flt(row[col_pct_blk] if col_pct_blk < len(row) else None)
        pct_of_attrition_raw = _flt(row[col_pct_att] if col_pct_att < len(row) else None)
        # Normalize: Excel stores % as decimal (0.76); we keep 0-100 range in DB
        pct_of_block     = round(pct_of_block_raw * 100, 1) if pct_of_block_raw is not None and pct_of_block_raw <= 1.5 else pct_of_block_raw
        pct_of_attrition = round(pct_of_attrition_raw * 100, 1) if pct_of_attrition_raw is not None and pct_of_attrition_raw <= 1.5 else pct_of_attrition_raw
        ota_rate         = _flt(row[col_ota]     if col_ota     < len(row) else None)

        if not pickup_by_night and total_rooms is None:
            continue  # empty row

        weekly_entries.append({
            'report_date':      report_date,
            'pickup_by_night':  json.dumps(pickup_by_night),
            'total_rooms':      total_rooms,
            'change_from_last': change_from_last,
            'pct_of_block':     pct_of_block,
            'pct_of_attrition': pct_of_attrition,
            'ota_rate':         ota_rate,
            'label':            label,
        })

    return {
        'config': {
            'tab_name':             tab_name,
            'organization':         organization,
            'event_name':           event_name,
            'hotel':                hotel,
            'hotel_contact':        hotel_contact,
            'hotel_contact_email':  hotel_contact_email,
            'group_contact':        group_contact,
            'booking_id':           booking_id,
            'cutoff_date':          cutoff_date,
            'attrition_pct':        attrition_pct,
            'contracted_block':     contracted_block,   # dict, not JSON string
        },
        'weekly_entries': weekly_entries,
    }


# ── PDF rooming list parser ───────────────────────────────────────────────────

# Matches a guest data line: Name(s) + 8-digit conf# + arr MM/DD/YY + dep MM/DD/YY
# followed by optional room type, status, persons, nights, rooms columns.
# Name group uses [ \t] (not \s) so it cannot cross a line boundary — this prevents
# "Block Code XYZ" section headers from being absorbed into the first guest's name.
# An optional leading room number (e.g. "1122 Larson,Thad") is consumed before the name.
_GUEST_RE = re.compile(
    r'^\d*[ \t]*'                                    # optional leading room number
    r'([A-Za-z][A-Za-z,\'\-\.\ \t]{1,40}?)[ \t]+'  # Last,First[,Title] — no newlines
    r'(\d{8})[ \t]+'                                 # 8-digit conf #
    r'(\d{2}/\d{2}/\d{2})[ \t]+'                    # arrival MM/DD/YY
    r'(\d{2}/\d{2}/\d{2})[ \t]+'                    # departure MM/DD/YY
    r'\S+[ \t]+'                                     # room type
    r'\d+[ \t]+'                                     # reservation status (numeric code)
    r'(\d+)[ \t]+'                                   # persons
    r'(\d+)[ \t]+'                                   # nights
    r'(\d+)',                                        # rooms
    re.MULTILINE
)

# Cambria/Choice Hotels "Group Reservation List" format:
#   Name  Account(9-12 digits)  Status(letter)  Arrival(M/DD/YYYY)  Departure  People  RoomType  Rate  Balance
# No "nights" column — computed from arrival/departure.
_GUEST_RE_CAMBRIA = re.compile(
    r'^([A-Za-z][A-Za-z,\'\-\.\ \t]{1,40}?)[ \t]+'  # Guest Name (Last, First)
    r'(\d{9,12})[ \t]+'                               # Account / conf# (9–12 digits)
    r'[A-Z][ \t]+'                                    # Status letter (R, C, etc.)
    r'(\d{1,2}/\d{2}/\d{4})[ \t]+'                   # Arrival   M/DD/YYYY
    r'(\d{1,2}/\d{2}/\d{4})[ \t]+'                   # Departure M/DD/YYYY
    r'(\d+)[ \t]+'                                    # People
    r'[A-Z]+[ \t]+'                                   # Room Type
    r'[\d,]+\.\d+[ \t]+'                              # Rate
    r'[\d,]+\.\d+',                                   # Balance
    re.MULTILINE
)

# Royal Sonesta "Group Rooming List" format:
#   Last,First  Conf#(9 digits)  Arrival(MM-DD-YY)  Departure(MM-DD-YY)  RoomType[+Status]  [Status]  Adl  Chl  Nts  Rms  ...
# Room type and status may be merged into one token (e.g. "U1KNRGC") or separate ("S1K GC").
# "Res. Notes:" continuation lines are ignored because they don't match this pattern.
_GUEST_RE_SONESTA = re.compile(
    r'^([A-Za-z][A-Za-z,\'\-\.\ \t]{1,40}?)[ \t]+'  # Last,First name
    r'(\d{9})[ \t]+'                                  # 9-digit conf#
    r'(\d{2}-\d{2}-\d{2})[ \t]+'                     # Arrival   MM-DD-YY (dashes)
    r'(\d{2}-\d{2}-\d{2})[ \t]+'                     # Departure MM-DD-YY (dashes)
    r'\S+[ \t]+'                                      # Room type token (may include status)
    r'(?:[A-Za-z]+[ \t]+)?'                           # Optional separate status code
    r'\d+[ \t]+'                                      # Adults
    r'\d+[ \t]+'                                      # Children
    r'(\d+)[ \t]+'                                    # Nights (explicit)
    r'(\d+)',                                          # Rooms
    re.MULTILINE
)

def _mddyyyy_to_iso(s):
    """Convert M/DD/YYYY or MM/DD/YYYY to YYYY-MM-DD."""
    parts = s.strip().split('/')
    m, d, y = int(parts[0]), int(parts[1]), int(parts[2])
    return f'{y}-{m:02d}-{d:02d}'

def _mmddyy_to_iso(s):
    """Convert MM/DD/YY to YYYY-MM-DD."""
    m, d, y = s.split('/')
    yr = 2000 + int(y) if int(y) < 50 else 1900 + int(y)
    return f'{yr}-{int(m):02d}-{int(d):02d}'

def _mmddyy_dashes_to_iso(s):
    """Convert MM-DD-YY (dashes) to YYYY-MM-DD."""
    m, d, y = s.split('-')
    yr = 2000 + int(y) if int(y) < 50 else 1900 + int(y)
    return f'{yr}-{int(m):02d}-{int(d):02d}'


def _ai_parse_rooming_list(all_text):
    """
    AI fallback: send raw PDF text to Claude and ask it to extract guest records.
    Used when no regex pattern recognizes the hotel's format.

    Returns a list of guest dicts (same shape as regex parsers), or [] if
    the API key is not configured or the call fails.
    """
    # Load API key from config.py, fall back to environment variable
    api_key = ''
    try:
        import importlib, sys
        # Force reload so edits to config.py take effect without restarting
        if 'config' in sys.modules:
            importlib.reload(sys.modules['config'])
        else:
            import config as _cfg
            sys.modules['config'] = _cfg
        api_key = sys.modules['config'].ANTHROPIC_API_KEY.strip()
    except Exception:
        pass
    if not api_key:
        import os
        api_key = os.environ.get('ANTHROPIC_API_KEY', '').strip()
    if not api_key:
        return [], None

    try:
        import anthropic
    except ImportError:
        return [], None

    prompt = (
        "You are parsing a hotel group rooming list exported from a property management system.\n"
        "Extract every individual guest reservation from the text below.\n\n"
        "Return ONLY a valid JSON array — no other text, no markdown fences, no explanation.\n"
        "Each element must have exactly these keys:\n"
        "  name       — guest name as it appears (e.g. \"Smith,John\" or \"John Smith\")\n"
        "  conf_no    — reservation/confirmation number (string)\n"
        "  arrival    — arrival date in YYYY-MM-DD format\n"
        "  departure  — departure date in YYYY-MM-DD format\n"
        "  nights     — integer number of nights\n"
        "  rooms      — integer number of rooms (default 1 if not shown)\n\n"
        "Rules:\n"
        "- Skip header rows, page headers/footers, summary/total rows, and note lines.\n"
        "- If 'nights' is not explicit, compute it as (departure - arrival) in days.\n"
        "- Convert all dates to YYYY-MM-DD regardless of the source format.\n"
        "- If a line has an obvious continuation (e.g. 'Res. Notes:'), skip it.\n\n"
        "Rooming list text:\n"
        + all_text[:12000]   # ~3k tokens; enough for 100+ guests
    )

    try:
        client = anthropic.Anthropic(api_key=api_key)
        message = client.messages.create(
            model='claude-opus-4-5',
            max_tokens=4096,
            messages=[{'role': 'user', 'content': prompt}]
        )
        response_text = message.content[0].text.strip()
        # Strip markdown code fences if the model wrapped the JSON
        response_text = re.sub(r'^```[a-z]*\s*', '', response_text)
        response_text = re.sub(r'\s*```$', '', response_text)

        guests_raw = json.loads(response_text)

        result = []
        for g in guests_raw:
            try:
                arr_str = str(g.get('arrival', '')).strip()
                dep_str = str(g.get('departure', '')).strip()
                nights  = int(g.get('nights', 0))
                if nights <= 0:
                    nights = (datetime.strptime(dep_str, '%Y-%m-%d') -
                              datetime.strptime(arr_str, '%Y-%m-%d')).days
                if nights <= 0:
                    continue
                result.append({
                    'name':      str(g.get('name', '')).strip(),
                    'conf_no':   str(g.get('conf_no', '')).strip(),
                    'arrival':   arr_str,
                    'departure': dep_str,
                    'nights':    nights,
                    'rooms':     int(g.get('rooms', 1)),
                })
            except Exception:
                continue

        return result, None

    except Exception as e:
        return [], str(e)

def parse_rooming_list_pdf(file_bytes):
    """
    Parse a hotel rooming list PDF (pdfplumber).

    Returns:
      {
        'guests':        [{name, conf_no, arrival, departure, nights, rooms}],
        'nights_by_date': {"YYYY-MM-DD": N, ...},   # rooms per night
        'total_guests':  N,
        'error':         str or None
      }
    """
    try:
        import pdfplumber
    except ImportError:
        return {'guests': [], 'nights_by_date': {}, 'total_guests': 0,
                'error': 'pdfplumber not installed'}

    try:
        guests = []
        all_text = ''

        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            for page in pdf.pages:
                text = page.extract_text() or ''
                all_text += '\n' + text

        for m in _GUEST_RE.finditer(all_text):
            name_raw  = m.group(1).strip().rstrip(',')
            conf_no   = m.group(2)
            arr_raw   = m.group(3)
            dep_raw   = m.group(4)
            nights    = int(m.group(6))
            rooms     = int(m.group(7))

            try:
                arrival    = _mmddyy_to_iso(arr_raw)
                departure  = _mmddyy_to_iso(dep_raw)
            except Exception:
                continue

            guests.append({
                'name':      name_raw,
                'conf_no':   conf_no,
                'arrival':   arrival,
                'departure': departure,
                'nights':    nights,
                'rooms':     rooms,
            })

        # ── Fallback: Cambria/Choice Hotels "Group Reservation List" format ──
        if not guests:
            for m in _GUEST_RE_CAMBRIA.finditer(all_text):
                name_raw = m.group(1).strip().rstrip(',')
                conf_no  = m.group(2)
                arr_raw  = m.group(3)
                dep_raw  = m.group(4)
                try:
                    arrival   = _mddyyyy_to_iso(arr_raw)
                    departure = _mddyyyy_to_iso(dep_raw)
                    nights    = (datetime.strptime(departure, '%Y-%m-%d') -
                                 datetime.strptime(arrival,   '%Y-%m-%d')).days
                except Exception:
                    continue
                if nights <= 0:
                    continue
                guests.append({
                    'name':      name_raw,
                    'conf_no':   conf_no,
                    'arrival':   arrival,
                    'departure': departure,
                    'nights':    nights,
                    'rooms':     1,
                })

        # ── Fallback: Royal Sonesta "Group Rooming List" format (MM-DD-YY dashes) ──
        if not guests:
            for m in _GUEST_RE_SONESTA.finditer(all_text):
                name_raw = m.group(1).strip().rstrip(',')
                conf_no  = m.group(2)
                arr_raw  = m.group(3)
                dep_raw  = m.group(4)
                nights   = int(m.group(5))
                rooms    = int(m.group(6))
                try:
                    arrival   = _mmddyy_dashes_to_iso(arr_raw)
                    departure = _mmddyy_dashes_to_iso(dep_raw)
                except Exception:
                    continue
                if nights <= 0:
                    continue
                guests.append({
                    'name':      name_raw,
                    'conf_no':   conf_no,
                    'arrival':   arrival,
                    'departure': departure,
                    'nights':    nights,
                    'rooms':     rooms,
                })

        # ── Final fallback: AI parser for unrecognized formats ──
        ai_parsed = False
        ai_error  = None
        if not guests:
            ai_guests, ai_error = _ai_parse_rooming_list(all_text)
            if ai_guests:
                guests    = ai_guests
                ai_parsed = True

        # Compute rooms per night from arrival/departure dates
        nights_by_date = {}
        for g in guests:
            try:
                arr = datetime.strptime(g['arrival'],   '%Y-%m-%d')
                dep = datetime.strptime(g['departure'], '%Y-%m-%d')
                cur = arr
                while cur < dep:
                    key = cur.strftime('%Y-%m-%d')
                    nights_by_date[key] = nights_by_date.get(key, 0) + g['rooms']
                    cur += timedelta(days=1)
            except Exception:
                continue

        return {
            'guests':         guests,
            'nights_by_date': nights_by_date,
            'total_guests':   len(guests),
            'ai_parsed':      ai_parsed,
            'ai_error':       ai_error,
            'error':          None,
        }

    except Exception as e:
        return {'guests': [], 'nights_by_date': {}, 'total_guests': 0,
                'ai_parsed': False, 'error': str(e)}


# ── CSV / XLS rooming list parser ────────────────────────────────────────────

def _parse_date_flexible(val):
    """Try multiple date formats and return YYYY-MM-DD, or None."""
    if val is None:
        return None
    s = str(val).strip()
    for fmt in ('%m/%d/%Y', '%m/%d/%y', '%Y-%m-%d', '%d-%b-%Y', '%d/%m/%Y',
                '%B %d, %Y', '%b %d, %Y'):
        try:
            from datetime import datetime as _dt
            return _dt.strptime(s, fmt).strftime('%Y-%m-%d')
        except ValueError:
            continue
    # Excel serial date (float/int)
    try:
        serial = float(s)
        from datetime import datetime as _dt, timedelta as _td
        base = _dt(1899, 12, 30)
        return (base + _td(days=serial)).strftime('%Y-%m-%d')
    except Exception:
        pass
    return None


def _col_index(headers, *candidates):
    """Return the index of the first matching header (case-insensitive), or None."""
    lower = [h.lower().strip() for h in headers]
    for c in candidates:
        for i, h in enumerate(lower):
            if c.lower() in h:
                return i
    return None


def _build_guest_result(guests):
    """
    Compute nights_by_date from a parsed guest list, deduplicating by
    Hotel Confirmation # so that two guests sharing a room count as 1 room-night.

    Returns dedup_summary: list of {conf_no, names, nights} for shared rooms.
    """
    # Group guests by conf_no (non-empty)
    from collections import OrderedDict
    conf_groups = OrderedDict()  # conf_no -> list of guests
    no_conf = []
    for g in guests:
        cn = (g.get('conf_no') or '').strip()
        if cn:
            conf_groups.setdefault(cn, []).append(g)
        else:
            no_conf.append(g)

    # Build dedup summary (conf_nos with >1 guest = shared room)
    dedup_summary = []
    for cn, grp in conf_groups.items():
        if len(grp) > 1:
            dedup_summary.append({
                'conf_no': cn,
                'count':   len(grp),
                'names':   ', '.join(g['name'] for g in grp),
                'arrival': grp[0].get('arrival', ''),
                'departure': grp[0].get('departure', ''),
            })

    # Compute nights_by_date using ONE room per unique conf_no
    nights_by_date = {}
    # Each unique conf_no = 1 room
    representative_guests = [grp[0] for grp in conf_groups.values()] + no_conf
    for g in representative_guests:
        try:
            arr = datetime.strptime(g['arrival'],   '%Y-%m-%d')
            dep = datetime.strptime(g['departure'], '%Y-%m-%d')
            cur = arr
            while cur < dep:
                key = cur.strftime('%Y-%m-%d')
                nights_by_date[key] = nights_by_date.get(key, 0) + 1
                cur += timedelta(days=1)
        except Exception:
            continue

    # Unique room count = unique conf_nos + guests with no conf_no
    unique_rooms = len(conf_groups) + len(no_conf)

    return {
        'guests':         guests,
        'nights_by_date': nights_by_date,
        'total_guests':   len(guests),
        'unique_rooms':   unique_rooms,
        'dedup_summary':  dedup_summary,
        'error':          None,
    }


def _parse_date_range(val):
    """
    Parse a combined date-range string like 'May 23, 2026 - May 29, 2026'
    Returns (arrival_iso, departure_iso) or (None, None).
    """
    if not val:
        return None, None
    # Try splitting on ' - ' or ' – ' (en-dash)
    for sep in (' - ', ' – ', '-'):
        parts = str(val).split(sep, 1)
        if len(parts) == 2:
            arr = _parse_date_flexible(parts[0].strip())
            dep = _parse_date_flexible(parts[1].strip())
            if arr and dep:
                return arr, dep
    return None, None


def parse_rooming_list_csv(file_bytes):
    """
    Parse a hotel rooming list CSV.

    Handles two date layouts:
      A) Separate arrival + departure columns
      B) Single combined column e.g. 'Guest Reservation Dates': 'May 23, 2026 - May 29, 2026'
    """
    import csv, io as _io
    try:
        text = file_bytes.decode('utf-8-sig', errors='replace')
        # Handle tab-separated files exported as .csv
        sample = text[:2000]
        delimiter = '\t' if sample.count('\t') > sample.count(',') else ','
        reader = csv.reader(_io.StringIO(text), delimiter=delimiter)
        rows = list(reader)
        if not rows:
            return {'guests': [], 'nights_by_date': {}, 'total_guests': 0,
                    'error': 'Empty CSV file'}

        # Find header row
        header_row_idx = 0
        headers = rows[0]
        for i, row in enumerate(rows[:10]):
            joined = ' '.join(row).lower()
            if any(k in joined for k in ('arrival', 'check-in', 'checkin', 'check in',
                                          'last name', 'first name', 'guest name', 'name',
                                          'reservation date', 'stay date')):
                headers = row
                header_row_idx = i
                break

        # Map columns
        i_last   = _col_index(headers, 'last name', 'last')
        i_first  = _col_index(headers, 'first name', 'first')
        i_name   = _col_index(headers, 'guest name', 'full name', 'name')
        i_arr    = _col_index(headers, 'arrival', 'check-in', 'checkin', 'check in',
                                       'arrive', 'arr date', 'arrival date')
        i_dep    = _col_index(headers, 'departure', 'check-out', 'checkout', 'check out',
                                       'depart', 'dep date', 'departure date')
        # Combined date range column (e.g. Passkey/Cvent export)
        i_range  = _col_index(headers, 'reservation date', 'stay date', 'guest reservation',
                                       'dates', 'date range', 'hotel date')
        i_conf   = _col_index(headers, 'hotel confirmation', 'confirmation #', 'hotel conf',
                                       'conf #', 'conf no', 'confirmation', 'reservation #',
                                       'passkey', 'booking ref', 'res id')
        i_rooms  = _col_index(headers, 'rooms', 'room count', 'qty', 'quantity')
        i_nights = _col_index(headers, 'nights', 'night count', 'duration', 'los')

        # Need either (arrival + departure) OR a combined date-range column
        has_split_dates = (i_arr is not None and i_dep is not None)
        has_range_date  = (i_range is not None)
        if not has_split_dates and not has_range_date:
            return {'guests': [], 'nights_by_date': {}, 'total_guests': 0,
                    'error': 'Could not find date columns. '
                             f'Headers found: {list(headers)}'}

        guests = []
        for row in rows[header_row_idx + 1:]:
            if not row or all(c.strip() == '' for c in row):
                continue
            def cell(idx):
                return row[idx].strip() if idx is not None and idx < len(row) else ''

            # Build name
            if i_last is not None and i_first is not None:
                name = f"{cell(i_last)}, {cell(i_first)}".strip(', ')
            elif i_name is not None:
                name = cell(i_name)
            else:
                name = 'Unknown'

            # Parse dates
            if has_split_dates:
                arrival   = _parse_date_flexible(cell(i_arr))
                departure = _parse_date_flexible(cell(i_dep))
            else:
                arrival, departure = _parse_date_range(cell(i_range))
            if not arrival or not departure:
                continue

            conf_no = cell(i_conf) if i_conf is not None else ''
            rooms   = int(cell(i_rooms)) if i_rooms is not None and cell(i_rooms).isdigit() else 1
            if i_nights is not None and cell(i_nights).isdigit():
                nights = int(cell(i_nights))
            else:
                try:
                    arr_d = datetime.strptime(arrival,   '%Y-%m-%d')
                    dep_d = datetime.strptime(departure, '%Y-%m-%d')
                    nights = (dep_d - arr_d).days
                except Exception:
                    nights = 0

            guests.append({
                'name':      name,
                'conf_no':   conf_no,
                'arrival':   arrival,
                'departure': departure,
                'nights':    nights,
                'rooms':     rooms,
            })

        return _build_guest_result(guests)

    except Exception as e:
        return {'guests': [], 'nights_by_date': {}, 'total_guests': 0,
                'error': str(e)}


def parse_rooming_list_xls(file_bytes, filename=''):
    """
    Parse a hotel rooming list XLS or XLSX file.

    Uses openpyxl for .xlsx and xlrd for .xls.
    """
    ext = (filename.rsplit('.', 1)[-1].lower()) if '.' in filename else 'xlsx'
    try:
        if ext == 'xls':
            try:
                import xlrd
            except ImportError:
                return {'guests': [], 'nights_by_date': {}, 'total_guests': 0,
                        'error': 'xlrd not installed — run: pip install xlrd'}
            import io as _io
            wb = xlrd.open_workbook(file_contents=file_bytes)
            ws = wb.sheet_by_index(0)
            rows = [[str(ws.cell_value(r, c)).strip() for c in range(ws.ncols)]
                    for r in range(ws.nrows)]
        else:
            import openpyxl
            import io as _io
            wb = openpyxl.load_workbook(_io.BytesIO(file_bytes), data_only=True)
            ws = wb.active
            rows = []
            for row in ws.iter_rows(values_only=True):
                rows.append([str(c).strip() if c is not None else '' for c in row])

        if not rows:
            return {'guests': [], 'nights_by_date': {}, 'total_guests': 0,
                    'error': 'Empty spreadsheet'}

        # Find header row
        header_row_idx = 0
        headers = rows[0]
        for i, row in enumerate(rows[:15]):
            joined = ' '.join(row).lower()
            if any(k in joined for k in ('arrival', 'check-in', 'checkin', 'check in',
                                          'last name', 'first name', 'guest name', 'name')):
                headers = row
                header_row_idx = i
                break

        # Map columns (same logic as CSV parser)
        i_last    = _col_index(headers, 'last name', 'last')
        i_first   = _col_index(headers, 'first name', 'first')
        i_name    = _col_index(headers, 'guest name', 'full name', 'name')
        i_arr     = _col_index(headers, 'arrival', 'check-in', 'checkin', 'check in',
                                        'arrive', 'arr date', 'arrival date')
        i_dep     = _col_index(headers, 'departure', 'check-out', 'checkout', 'check out',
                                        'depart', 'dep date', 'departure date')
        i_conf    = _col_index(headers, 'conf', 'confirmation', 'reservation', 'res no',
                                        'booking ref', 'res #', 'res id')
        i_rooms   = _col_index(headers, 'rooms', 'room count', 'qty', 'quantity')
        i_nights  = _col_index(headers, 'nights', 'night count', 'duration', 'los',
                                        'length of stay')

        if i_arr is None or i_dep is None:
            return {'guests': [], 'nights_by_date': {}, 'total_guests': 0,
                    'error': 'Could not find arrival/departure columns. '
                             f'Headers found: {headers}'}

        guests = []
        for row in rows[header_row_idx + 1:]:
            if not row or all(c == '' or c == 'None' for c in row):
                continue
            def cell(i):
                return row[i].strip() if i is not None and i < len(row) else ''

            if i_last is not None and i_first is not None:
                name = f"{cell(i_last)}, {cell(i_first)}".strip(', ')
            elif i_name is not None:
                name = cell(i_name)
            else:
                name = 'Unknown'

            arrival   = _parse_date_flexible(cell(i_arr))
            departure = _parse_date_flexible(cell(i_dep))
            if not arrival or not departure:
                continue

            conf_no = cell(i_conf) if i_conf is not None else ''
            rooms   = int(cell(i_rooms)) if i_rooms is not None and cell(i_rooms).isdigit() else 1
            if i_nights is not None and cell(i_nights).isdigit():
                nights = int(cell(i_nights))
            else:
                try:
                    arr_d = datetime.strptime(arrival,   '%Y-%m-%d')
                    dep_d = datetime.strptime(departure, '%Y-%m-%d')
                    nights = (dep_d - arr_d).days
                except Exception:
                    nights = 0

            guests.append({
                'name':      name,
                'conf_no':   conf_no,
                'arrival':   arrival,
                'departure': departure,
                'nights':    nights,
                'rooms':     rooms,
            })

        return _build_guest_result(guests)

    except Exception as e:
        return {'guests': [], 'nights_by_date': {}, 'total_guests': 0,
                'error': str(e)}


# ── Reconciliation ────────────────────────────────────────────────────────────

def reconcile(pickup_by_night, rooming_nights_by_date):
    """
    Compare the entered weekly pickup (dict date→rooms) against the rooming
    list extract (dict date→rooms).

    Returns:
      {'status': 'match'|'discrepancy', 'notes': str}
    """
    diffs = []
    all_dates = sorted(set(list(pickup_by_night.keys()) +
                           list(rooming_nights_by_date.keys())))
    for date in all_dates:
        p = pickup_by_night.get(date)
        r = rooming_nights_by_date.get(date, 0)
        if p is None:
            continue   # night not in contracted block
        if p != r:
            diffs.append(
                f"{date}: pickup says {p}, rooming list shows {r} "
                f"({'over' if r > p else 'under'} by {abs(r - p)})"
            )

    if diffs:
        return {'status': 'discrepancy', 'notes': '\n'.join(diffs)}
    return {'status': 'match', 'notes': 'Rooming list matches pickup report.'}


# ── Excel export row ──────────────────────────────────────────────────────────

def format_export_row(config_row, weekly_row):
    """
    Return a tab-separated string matching the pickup spreadsheet row layout:
      Date | night1 | night2 | ... | total | change | %total | %attrition | webrate

    config_row  — sqlite3.Row from pickup_config
    weekly_row  — sqlite3.Row from pickup_weekly
    """
    try:
        block   = json.loads(config_row['contracted_block'] or '{}')
        pickup  = json.loads(weekly_row['pickup_by_night']  or '{}')
        dates   = sorted(block.keys())

        parts = [weekly_row['report_date']]
        for d in dates:
            v = pickup.get(d)
            parts.append(str(v) if v is not None else '')

        parts.append(str(weekly_row['total_rooms']      or ''))
        parts.append(str(weekly_row['change_from_last'] or ''))

        pob = weekly_row['pct_of_block']
        parts.append(f'{pob:.4f}' if pob is not None else '')

        poa = weekly_row['pct_of_attrition']
        parts.append(f'{poa:.4f}' if poa is not None else '')

        ota = weekly_row['ota_rate']
        parts.append(str(int(ota)) if ota is not None else '')

        return '\t'.join(parts)
    except Exception as e:
        return f'Error: {e}'


# ── Email generators ──────────────────────────────────────────────────────────

_SIGNATURE = ""

def _fmt_date(iso):
    """'2026-01-13' → 'January 13, 2026'"""
    try:
        return datetime.strptime(iso, '%Y-%m-%d').strftime('%B %d, %Y')
    except Exception:
        return iso

def _fmt_short(iso):
    """'2026-01-13' → 'MM/DD/YYYY'"""
    try:
        return datetime.strptime(str(iso)[:10], '%Y-%m-%d').strftime('%m/%d/%Y')
    except Exception:
        return iso or ''

def _build_cc(config_row):
    """Return comma-separated CC email string. Handles both old (str) and new (dict) formats."""
    try:
        entries = json.loads(config_row['cc_emails'] or '[]')
        addrs = []
        for e in entries:
            if isinstance(e, dict):
                addr = (e.get('email') or '').strip()
            else:
                addr = str(e).strip()
            if addr:
                addrs.append(addr)
        return ', '.join(addrs)
    except Exception:
        return ''


def _build_cc_recipients(config_row):
    """Return list of {'name': str, 'email': str} dicts for all CC contacts."""
    try:
        entries = json.loads(config_row['cc_emails'] or '[]')
        result = []
        for e in entries:
            if isinstance(e, dict):
                addr = (e.get('email') or '').strip()
                name = (e.get('name')  or '').strip()
            else:
                addr = str(e).strip()
                name = ''
            if addr:
                result.append({'name': name, 'email': addr})
        return result
    except Exception:
        return []


def build_hotel_email(config_row):
    """
    Build the hotel pickup-request email.
    Returns {'to': str, 'subject': str, 'body': str}.
    """
    try:
        block  = json.loads(config_row['contracted_block'] or '{}')
        dates  = sorted(block.keys())

        if dates:
            pre  = int(config_row['shoulder_pre']  or 3)
            post = int(config_row['shoulder_post'] or 3)
            first = datetime.strptime(dates[0], '%Y-%m-%d')
            last  = datetime.strptime(dates[-1], '%Y-%m-%d')
            shoulder_start = (first - timedelta(days=pre)).strftime('%B %d, %Y')
            shoulder_end   = (last  + timedelta(days=post)).strftime('%B %d, %Y')
            date_range = (f"{_fmt_date(dates[0])} – {_fmt_date(dates[-1])}"
                          if len(dates) > 1 else _fmt_date(dates[0]))
            block_lines = '\n'.join(
                f"  {_fmt_date(d)}: {block[d]} rooms contracted"
                for d in dates
            )
        else:
            shoulder_start = shoulder_end = date_range = ''
            block_lines = '  (no contracted block dates entered)'

        org   = config_row['organization'] or ''
        hotel = config_row['hotel'] or ''
        to    = config_row['hotel_contact_email'] or ''

        # Extract first name from hotel_contact for salutation
        _hc = (config_row['hotel_contact'] or '').strip()
        if ',' in _hc:
            _hc_first = _hc.split(',', 1)[1].strip().split()[0]
        elif _hc:
            _hc_first = _hc.split()[0]
        else:
            _hc_first = 'Team'

        subject = f"Pickup Request – {org} / {hotel} / {date_range}"

        body = (
            f"Hello {_hc_first},\n\n"
            f"I hope you are doing well.  Can you please send me the current group pickup report "
            f"and group rooming list for the following group:\n\n"
            f"  Organization: {org}\n"
            f"  Event: {config_row['event_name'] or ''}\n"
            f"  Hotel: {hotel}\n"
            f"  Meeting Dates: {date_range}\n\n"
            f"Contracted Block:\n{block_lines}\n\n"
            f"Please include shoulder nights ({shoulder_start} – {shoulder_end}) "
            f"in both the pickup report and rooming list.\n\n"
            f"Please send:\n"
            f"  1. Day-by-day pickup report\n"
            f"  2. Group rooming list\n\n"
            f"\n"
            f""
        )

        cc = _build_cc(config_row)
        return {'to': to, 'cc': cc, 'subject': subject, 'body': body}
    except Exception as e:
        return {'to': '', 'cc': '', 'subject': 'Pickup Request', 'body': f'Error generating email: {e}'}


def hotel_request_mailto(config_row):
    """Return a mailto: URL for the hotel pickup request (kept for backwards compat)."""
    e = build_hotel_email(config_row)
    return f"mailto:{quote(e['to'])}?subject={quote(e['subject'])}&body={quote(e['body'])}"


def build_client_email(config_row, weekly_row, rl_status=None, weekly_list=None):
    """
    Build the client pickup-summary email with a spreadsheet-style grid.
    Returns {'to': str, 'subject': str, 'body': str}.
    weekly_row may be None if no entries exist yet.
    weekly_list is the full list of weekly entry dicts (most-recent first).
    """
    try:
        block      = json.loads(config_row['contracted_block'] or '{}')
        dates      = sorted(block.keys())
        org        = config_row['organization'] or ''
        event_name = config_row['event_name'] or ''
        hotel      = config_row['hotel'] or ''
        to         = config_row['group_contact_email'] or ''

        # Extract first name from group_contact for salutation
        # Handles "First Last", "Last, First", or a single name
        _gc = (config_row['group_contact'] or '').strip()
        if ',' in _gc:
            # "Last, First" format — take everything after the comma
            _first = _gc.split(',', 1)[1].strip().split()[0]
        elif _gc:
            _first = _gc.split()[0]
        else:
            _first = 'Team'

        if not weekly_row:
            subject = f"Pickup Update – {event_name or org} | {hotel}"
            body = (
                f"Hello {_first},\n\n"
                f"I wanted to touch base regarding the upcoming group at {hotel}.\n"
                f"No pickup data has been recorded yet — I will follow up once the "
                f"first report is available.\n\n"
                f""
            )
            return {'to': to, 'subject': subject, 'body': body}

        pickup    = json.loads(weekly_row['pickup_by_night'] or '{}')
        report_dt = weekly_row['report_date'] or ''
        total_blk = sum(block.get(d, 0) for d in dates)
        ota       = weekly_row['ota_rate']
        c_rate    = config_row['contracted_rate']
        _atr_pct_dec = config_row['attrition_pct'] or 0  # decimal e.g. 0.80
        _atr_rooms   = total_blk * float(_atr_pct_dec) if total_blk else 0

        # Always recompute totals/pcts from pickup_by_night (source of truth)
        def _recompute(w):
            """Return a plain dict with corrected total_rooms / pct_of_block /
               pct_of_attrition.  change_from_last is filled in a second pass."""
            pbn = json.loads(w['pickup_by_night'] or '{}')
            t = sum(v for v in pbn.values() if v is not None and v != '')
            pob_ = round(t / total_blk * 100, 1) if total_blk else None
            poa_ = round(t / _atr_rooms * 100, 1) if _atr_rooms else None
            d = dict(w)
            d['total_rooms']      = t
            d['pct_of_block']     = pob_
            d['pct_of_attrition'] = poa_
            return d

        wr = _recompute(weekly_row)
        total = wr['total_rooms']
        pob   = wr['pct_of_block']
        poa   = wr['pct_of_attrition']

        pob_str = f"{pob:.1f}%" if pob is not None else 'N/A'
        poa_str = f"{poa:.1f}%" if poa is not None else 'N/A'

        # Attrition balance — rooms still needed to reach the attrition floor
        if _atr_pct_dec and total_blk:
            _atr_target  = round(total_blk * float(_atr_pct_dec))
            _atr_balance = _atr_target - total
            if _atr_balance > 0:
                _atr_note = f"  ({_atr_balance} room-nights still needed to reach attrition target)"
            else:
                _atr_note = f"  (attrition target reached ✓)"
        else:
            _atr_note = ""

        if pob is not None and pob >= 80:
            status_line = "Your group is ON PACE with the contracted block."
        elif pob is not None and pob >= 60:
            status_line = "Your group is slightly BELOW PACE — monitoring closely."
        else:
            status_line = ""

        # ── Build label-per-line format (works in any font, no alignment needed) ──
        def fmt_night(iso):
            try:
                dt = datetime.strptime(iso, '%Y-%m-%d')
                return dt.strftime('%a %m/%d/%y')
            except Exception:
                return iso

        grid = 'NIGHT-BY-NIGHT PICKUP\n'
        for d in dates:
            b = block.get(d, 0) or 0
            p = pickup.get(d, 0) or 0
            night = fmt_night(d)
            if not b:
                grid += f"  {night}:  {p} rooms  (shoulder)\n"
            else:
                diff = p - b
                diff_s = f'+{diff}' if diff > 0 else str(diff)
                grid += f"  {night}:  Block {b} / Pickup {p}  ({diff_s})\n"
        total_diff = total - total_blk
        total_diff_s = f'+{total_diff}' if total_diff > 0 else str(total_diff)
        grid += f"  Total:  Block {total_blk:,} / Pickup {total:,}  ({total_diff_s})\n"
        grid += '\n'

        # Weekly history — recompute all entries from pickup_by_night then
        # fill change_from_last correctly (oldest → newest pass)
        raw_entries = weekly_list or [dict(weekly_row)]
        # Sort oldest-first for WoW calculation, then flip back
        recomputed = [_recompute(w) for w in reversed(list(raw_entries))]
        prev_t = None
        for e in recomputed:
            e['change_from_last'] = (e['total_rooms'] - prev_t) if prev_t is not None else None
            prev_t = e['total_rooms']
        recomputed.reverse()   # back to most-recent-first
        entries = recomputed
        history_lines = []
        for w in entries:
            w_total  = w['total_rooms']      if isinstance(w, dict) else w.get('total_rooms')
            w_change = w['change_from_last'] if isinstance(w, dict) else w.get('change_from_last')
            w_pob    = w['pct_of_block']     if isinstance(w, dict) else w.get('pct_of_block')
            w_poa    = w['pct_of_attrition'] if isinstance(w, dict) else w.get('pct_of_attrition')
            w_date   = (w['report_date']     if isinstance(w, dict) else w.get('report_date')) or ''
            w_label  = w.get('label', '') if isinstance(w, dict) else ''
            try:
                w_date_fmt = datetime.strptime(w_date, '%Y-%m-%d').strftime('%m/%d/%y')
            except Exception:
                w_date_fmt = w_date
            chg_str = (f'+{w_change}' if w_change and w_change > 0
                       else str(w_change)) if w_change is not None else '—'
            pob_s = f'{w_pob:.1f}%'  if w_pob  is not None else '—'
            poa_s = f'{w_poa:.1f}%'  if w_poa  is not None else '—'
            tot_s = f'{w_total:,}'   if w_total is not None else '—'
            lbl   = f' ({w_label})'  if w_label else ''
            history_lines.append(
                f"  {w_date_fmt}{lbl}:  {tot_s} rooms  ({chg_str})  |  {pob_s} of block  |  {poa_s} of attrition\n"
            )

        history_block = 'WEEKLY HISTORY\n' + ''.join(history_lines)

        # OTA note
        ota_note = ''
        if ota and c_rate:
            diff = float(ota) - float(c_rate)
            if diff < 0:
                ota_note = (
                    f"\nOTA RATE ALERT: The hotel's current online rate "
                    f"(${float(ota):,.2f}) is BELOW the contracted group rate "
                    f"(${float(c_rate):,.2f}) by ${abs(diff):,.2f}. "
                    f"I will follow up with the hotel immediately."
                )
            else:
                ota_note = (
                    f"\nOTA Rate Check: Online rate ${float(ota):,.2f} is above "
                    f"the contracted group rate (${float(c_rate):,.2f})."
                )
        elif ota:
            ota_note = f"\n  Current OTA rate: ${float(ota):,.2f}"

        # Rooming list note
        rl_note = ''
        if rl_status == 'match':
            rl_note = '\nRooming list verified — matches pickup report.'
        elif rl_status == 'discrepancy':
            rl_note = '\nRooming list discrepancy noted — following up with hotel.'

        subject = f"Pickup Update – {event_name or org} | {hotel} | As of {_fmt_short(report_dt)}"

        try:
            # Use the recomputed entries list — first entry is most recent
            _wow = entries[0]['change_from_last'] if entries else None
        except Exception:
            _wow = None
        wow_s = (f'+{_wow}' if _wow and _wow > 0 else str(_wow)) if _wow is not None else '—'

        body = (
            f"Hello {_first},\n\n"
            f"Here is your weekly pickup update for {org} at {hotel}.\n"
            f"Report date: {_fmt_short(report_dt)}   |   Cut-off: {_fmt_short(config_row['cutoff_date']) if config_row['cutoff_date'] else 'N/A'}\n\n"
            f"SUMMARY\n"
            f"  Total Rooms:      {total:,} of {total_blk:,}\n"
            f"  % of Block:       {pob_str}\n"
            f"  % of Attrition:   {poa_str}{_atr_note}\n"
            f"  Week-over-Week:   {wow_s}\n\n"
            + (f"{status_line}\n\n" if status_line else "")
            + f"{grid}\n"
            + f"{history_block}\n"
            + (f"{ota_note}\n" if ota_note else "")
            + (f"{rl_note}\n" if rl_note else "")
            + "\nPlease feel free to reach out with any questions.\n\n"
        )

        # ── HTML version — spreadsheet-style table matching weekly pickup history ─
        F   = 'font-family:Calibri,Arial,sans-serif; font-size:10pt;'
        B   = f'{F} border:1px solid #aaa; padding:3px 6px;'
        def td(extra=''):  return f'style="{B} {extra}"'
        def th(extra=''):  return f'style="{B} background:#dce6f1; font-weight:bold; {extra}"'
        def tf(extra=''):  return f'style="{B} background:#f2f2f2; font-weight:bold; {extra}"'

        def diff_color(d):
            return 'color:green;' if d >= 0 else 'color:#cc0000;'

        # Build the single combined spreadsheet table:
        # Rows: header (dates) | Day | Block | [weekly entries, newest first] | Remaining
        # Columns: label | date1 | date2 | ... | Total | WoW | % Block | % Attr
        n = len(dates)
        summary_hdrs = ['Total', 'WoW', '% Block', '% Attr']
        TS = f'style="border-collapse:collapse; {F}"'

        # Header row — dates
        date_headers = ''.join(
            f'<th {th("text-align:center;")}>'
            + (datetime.strptime(d, "%Y-%m-%d").strftime("%m/%d") if len(d) == 10 else d)
            + '</th>'
            for d in dates
        )
        summary_headers = ''.join(f'<th {th("text-align:right;")}>{h}</th>' for h in summary_hdrs)
        header_row = f'<tr><th {th()}></th>{date_headers}{summary_headers}</tr>'

        # Day-of-week row
        day_cells = ''.join(
            f'<td {td("text-align:center; color:#555;")}>'
            + (datetime.strptime(d, "%Y-%m-%d").strftime("%a") if len(d) == 10 else '')
            + '</td>'
            for d in dates
        )
        day_row = f'<tr><td {td("color:#555;")}><i>Day</i></td>{day_cells}<td {td()} colspan="{len(summary_hdrs)}"></td></tr>'

        # Block row
        block_cells = ''.join(
            f'<td {td("text-align:right;")}>{block.get(d,0) or "—"}</td>'
            for d in dates
        )
        block_row = (
            f'<tr><td {tf()}>Block</td>{block_cells}'
            f'<td {tf("text-align:right;")}>{total_blk:,}</td>'
            f'<td {tf()} colspan="{len(summary_hdrs)-1}"></td></tr>'
        )

        # Weekly data rows (entries already ordered most-recent first)
        data_rows = []
        for w in entries:
            w_pickup = json.loads(w['pickup_by_night'] if isinstance(w, dict)
                                  else w.get('pickup_by_night') or '{}')
            w_total  = w['total_rooms']      if isinstance(w, dict) else w.get('total_rooms')
            w_change = w['change_from_last'] if isinstance(w, dict) else w.get('change_from_last')
            w_pob    = w['pct_of_block']     if isinstance(w, dict) else w.get('pct_of_block')
            w_poa    = w['pct_of_attrition'] if isinstance(w, dict) else w.get('pct_of_attrition')
            w_date   = (w['report_date']     if isinstance(w, dict) else w.get('report_date')) or ''
            w_label  = w.get('label', '')    if isinstance(w, dict) else ''
            try:
                w_date_fmt = datetime.strptime(w_date, '%Y-%m-%d').strftime('%m/%d/%y')
            except Exception:
                w_date_fmt = w_date
            lbl = f' ({w_label})' if w_label else ''
            chg_str = (f'+{w_change}' if w_change and w_change > 0 else str(w_change)) if w_change is not None else '—'
            pob_s   = f'{w_pob:.1f}%' if w_pob  is not None else '—'
            poa_s   = f'{w_poa:.1f}%' if w_poa  is not None else '—'
            tot_s   = f'{w_total:,}'  if w_total is not None else '—'
            chg_dc  = diff_color(w_change) if w_change is not None else ''
            night_cells = ''.join(
                f'<td {td("text-align:right;")}>{w_pickup.get(d, "—")}</td>'
                for d in dates
            )
            data_rows.append(
                f'<tr>'
                f'<td {td()}>{w_date_fmt}{lbl}</td>'
                + night_cells +
                f'<td {td("text-align:right; font-weight:bold;")}>{tot_s}</td>'
                f'<td style="{B} text-align:right; {chg_dc}">{chg_str}</td>'
                f'<td {td("text-align:right;")}>{pob_s}</td>'
                f'<td {td("text-align:right;")}>{poa_s}</td>'
                f'</tr>'
            )

        # Remaining row (block − latest pickup, per night)
        rem_cells = ''.join(
            f'<td {td("text-align:right;")}>{(block.get(d,0) or 0) - (pickup.get(d) or 0)}</td>'
            for d in dates
        )
        rem_total = total_blk - total
        rem_dc = diff_color(-rem_total)  # negative remaining = over block = green
        rem_row = (
            f'<tr><td {td("color:#555;")}><i>Remaining</i></td>{rem_cells}'
            f'<td style="{B} text-align:right; font-style:italic; {rem_dc}">{rem_total:+,}</td>'
            f'<td {td()} colspan="{len(summary_hdrs)-1}"></td></tr>'
        )

        pickup_table = (
            f'<table {TS} border="1" cellpadding="0" cellspacing="0">'
            f'<thead>{header_row}{day_row}</thead>'
            f'<tbody>{block_row}{"".join(data_rows)}{rem_row}</tbody>'
            f'</table>'
        )

        ota_html = ''
        if ota_note:
            color = 'color:#cc0000; font-weight:bold;' if 'ALERT' in ota_note else ''
            ota_html = f'<p style="{F} {color}">{ota_note.strip()}</p>'
        rl_html = f'<p style="{F}">{rl_note.strip()}</p>' if rl_note else ''
        status_html = (f'<p style="{F} font-weight:bold;">{status_line}</p>'
                       if status_line else '')
        P = f'style="{F} margin:6px 0;"'

        html_body = (
            f'<p {P}>Hello {_first},</p>'
            f'<p {P}>Here is your weekly pickup update for <b>{org}</b> at <b>{hotel}</b>.<br>'
            f'Report date: {report_dt}&nbsp;&nbsp;|&nbsp;&nbsp;Cut-off: {config_row["cutoff_date"] or "N/A"}</p>'
            f'<p {P}><b>SUMMARY</b><br>'
            f'Total Rooms: <b>{total:,}</b> of {total_blk:,}<br>'
            f'% of Block: <b>{pob_str}</b><br>'
            f'% of Attrition: <b>{poa_str}</b>{_atr_note}<br>'
            f'Week-over-Week: <b>{wow_s}</b></p>'
            + status_html
            + f'<p {P}><b>WEEKLY PICKUP REPORT</b></p>'
            + pickup_table
            + ota_html
            + rl_html
            + f'<p {P}>Please feel free to reach out with any questions.</p>'
        )

        cc = _build_cc(config_row)
        return {'to': to, 'cc': cc, 'subject': subject, 'body': body, 'html_body': html_body}
    except Exception as e:
        return {'to': '', 'cc': '', 'subject': 'Pickup Update', 'body': f'Error generating email: {e}'}


def client_summary_mailto(config_row, weekly_row, rl_status=None):
    """Return a mailto: URL for the client pickup summary (kept for backwards compat)."""
    e = build_client_email(config_row, weekly_row, rl_status)
    return f"mailto:{quote(e['to'])}?subject={quote(e['subject'])}&body={quote(e['body'])}"


# ── Contract Template Parsing ─────────────────────────────────────────────────

_OPTIONAL_SECTION_KEYWORDS = [
    'meeting room', 'meeting space', 'food', 'beverage', 'f&b', 'catering',
    'exhibit', 'banquet', 'function space', 'audio', 'visual', 'av ',
]


def extract_template_metadata(file_bytes):
    """
    Parse a Word contract template (.docx or .docm) and extract:
      - merge_fields: list of MERGEFIELD names found in XML
      - sections: list of {name, index, optional} dicts
    """
    import zipfile
    try:
        from lxml import etree as _et
    except ImportError:
        return {'merge_fields': [], 'sections': []}

    try:
        with zipfile.ZipFile(io.BytesIO(file_bytes)) as z:
            names = z.namelist()
            doc_name = 'word/document.xml' if 'word/document.xml' in names else next(
                (n for n in names if 'document' in n and n.endswith('.xml')), None)
            if doc_name is None:
                return {'merge_fields': [], 'sections': []}
            xml_content = z.read(doc_name).decode('utf-8', errors='replace')
    except Exception:
        return {'merge_fields': [], 'sections': []}

    fields = re.findall(r'MERGEFIELD\s+"?([A-Za-z_][A-Za-z0-9_ ]*)"?', xml_content)
    fields += re.findall(r'MERGEFIELD\s+([A-Za-z_][A-Za-z0-9_]+)\s', xml_content)
    merge_fields = sorted(set(f.strip().replace(' ', '_') for f in fields if f.strip()))

    sections = []
    idx = 0
    try:
        root = _et.fromstring(xml_content.encode('utf-8'))
        ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        for para in root.findall('.//w:p', ns):
            runs = para.findall('.//w:r', ns)
            text = ''.join(r.findtext('w:t', '', ns) for r in runs).strip()
            if not text or len(text) < 4:
                continue
            pPr = para.find('w:pPr', ns)
            style_val = ''
            if pPr is not None:
                pStyle = pPr.find('w:pStyle', ns)
                if pStyle is not None:
                    style_val = pStyle.get(
                        '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '')
            is_heading_style = 'heading' in style_val.lower()
            is_bold = len(para.findall('.//w:b', ns)) > 0
            leading = text.split(':')[0].strip()
            all_words = leading.split()
            consecutive_caps = 0
            for w in all_words:
                alpha = re.sub(r'[^A-Za-z]', '', w)
                if alpha and len(alpha) >= 2 and alpha.isupper():
                    consecutive_caps += 1
                else:
                    break
            is_caps_heading = consecutive_caps >= 2 and is_bold
            if is_heading_style or is_caps_heading:
                section_name = leading if len(leading) >= 4 else text
                skip_phrases = ['in witness whereof', 'hotel use only', 'undersigned expressly']
                if re.match(r'^[X☐☑✓]\s*[-–]', section_name):
                    continue
                if any(sp in section_name.lower() for sp in skip_phrases):
                    continue
                tl = section_name.lower()
                optional = any(kw in tl for kw in _OPTIONAL_SECTION_KEYWORDS)
                sections.append({'name': section_name, 'index': idx, 'optional': optional})
                idx += 1
    except Exception:
        pass

    return {'merge_fields': merge_fields, 'sections': sections}


# ── Cvent RFP Document Parser ─────────────────────────────────────────────────

def parse_rfp_docx(file_bytes):
    """
    Parse a Cvent RFP Word document (.docx) and extract key fields.
    Returns a dict with any of these keys populated (None if not found):
        rfp_name, rfp_code, event_name, client_org,
        start_date, end_date, response_due_date, decision_due_date,
        total_attendees, f_and_b_budget, total_room_nights, peak_rooms
    """
    import io as _io
    import re as _re
    import zipfile
    from lxml import etree as _et
    from datetime import datetime as _dt

    try:
        with zipfile.ZipFile(_io.BytesIO(file_bytes)) as z:
            xml = z.read('word/document.xml')
    except Exception:
        return {}

    root = _et.fromstring(xml)
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    lines = []
    for p in root.findall('.//w:p', ns):
        line = ''.join(t.text or '' for t in p.findall('.//w:t', ns)).strip()
        if line:
            lines.append(line)

    def _after(label_keywords, max_gap=2):
        for i, l in enumerate(lines):
            ll = l.lower()
            if all(kw.lower() in ll for kw in label_keywords):
                for j in range(1, max_gap + 1):
                    if i + j < len(lines):
                        v = lines[i + j].strip()
                        if v:
                            return v
        return None

    def _parse_date(s):
        if not s:
            return None
        s = _re.sub(r'^[A-Za-z]{3},\s*', '', s.strip())
        for fmt in ('%b %d, %Y', '%B %d, %Y', '%m/%d/%Y', '%Y-%m-%d'):
            try:
                return _dt.strptime(s.strip(), fmt).strftime('%Y-%m-%d')
            except ValueError:
                pass
        return None

    def _parse_date_range(s):
        if not s:
            return None, None
        parts = _re.split(r'\s*[-–]\s*', s, maxsplit=1)
        return _parse_date(parts[0]), _parse_date(parts[1]) if len(parts) > 1 else None

    def _parse_currency(s):
        if not s:
            return None
        m = _re.search(r'[\d,]+\.?\d*', s)
        return float(m.group().replace(',', '')) if m else None

    def _parse_int(s):
        if not s:
            return None
        m = _re.search(r'\d[\d,]*', s)
        return int(m.group().replace(',', '')) if m else None

    rfp_name          = _after(['RFP Name'])
    rfp_code          = _after(['RFP Code'])
    response_due      = _parse_date(_after(['Response Due Date']))
    decision_due      = _parse_date(_after(['Decision Due Date']))
    total_attendees   = _parse_int(_after(['Total Attendees']))
    f_and_b_budget    = _parse_currency(_after(['Food and Beverage Budget']))
    total_room_nights = _parse_int(_after(['Total Room Nights']))
    peak_rooms        = _parse_int(_after(['Peak Room Nights']))
    client_org        = _after(['Organization Name'])
    start_date, end_date = _parse_date_range(_after(['Event Dates']))

    return {
        'rfp_name':          rfp_name,
        'rfp_code':          rfp_code,
        'event_name':        rfp_name,
        'client_org':        client_org,
        'response_due_date': response_due,
        'decision_due_date': decision_due,
        'total_attendees':   total_attendees,
        'f_and_b_budget':    f_and_b_budget,
        'total_room_nights': total_room_nights,
        'peak_rooms':        peak_rooms,
        'start_date':        start_date,
        'end_date':          end_date,
    }


# ── Cvent CRF (Consolidated Response Form) Excel Parser ──────────────────────

def parse_crf_excel(file_bytes):
    """
    Parse a Cvent Consolidated Response Form (CRF) Excel file.
    Returns:
        {'rfp_meta': {event_name, response_due_date, decision_due_date},
         'hotels': [{hotel_name, city, state, status, proposed_rate,
                     f_and_b_minimum, commission_pct, attrition_pct,
                     cutoff_days, concessions, notes, contact_name,
                     contact_email, contact_phone, contact_title, crf_row_data}]}
    """
    import io as _io
    import re as _re
    import json as _json
    import openpyxl

    wb = openpyxl.load_workbook(_io.BytesIO(file_bytes), read_only=True, data_only=True)

    def _sv(v):
        if v is None:
            return ''
        s = str(v).replace('_x000d_', ' ').replace('\r', ' ')
        return _re.sub(r'\s+', ' ', s).strip()

    def _parse_currency(val):
        s = _sv(val)
        if not s or s.lower() in ('n/a', 'na', '—', '-'):
            return None
        cleaned = _re.sub(r'[USD$,]', '', s)
        m = _re.search(r'[\d]+\.?\d*', cleaned)
        return float(m.group()) if m else None

    def _parse_commission(val):
        s = _sv(val)
        m = _re.search(r'(\d+(?:\.\d+)?)\s*%', s)
        return float(m.group(1)) / 100.0 if m else None

    def _parse_address(val):
        if val is None:
            return None, None
        raw = str(val).replace('_x000d_', '\n').replace('\r\n', '\n').replace('\r', '\n')
        lines = [ln.strip() for ln in raw.split('\n') if ln.strip() and ln.strip().upper() != 'USA']
        city_state_line = None
        for line in lines:
            if _re.search(r'[A-Za-z].+,\s*[A-Za-z]', line):
                city_state_line = line
                break
        if not city_state_line:
            city_state_line = ' '.join(lines)
        parts = [p.strip() for p in city_state_line.split(',')]
        if len(parts) >= 2:
            city  = parts[0].strip()
            state = _re.sub(r'\s+[\d\-]+\s*$', '', parts[1].strip()).strip()
            state = _re.sub(r'\s*(USA|US|United States)\s*$', '', state, flags=_re.I).strip()
            return city, state
        return None, None

    def _map_status(val):
        s = _sv(val).lower()
        if 'submitted' in s:
            return 'proposal_received'
        if 'turned down' in s or 'withdrawn' in s or 'declined' in s:
            return 'declined'
        return 'pending'

    # ── RFP metadata sheet ────────────────────────────────────────────────────
    rfp_meta = {}
    if 'RFP' in wb.sheetnames:
        rfp_ws = wb['RFP']
        rfp_rows = list(rfp_ws.rows)
        for row in rfp_rows[:25]:
            vals = [c.value for c in row]
            if len(vals) < 2:
                continue
            label = _sv(vals[0]).lower()
            val = next((_sv(v) for v in vals[1:4] if _sv(v)), '')
            if ('response' in label and 'due' in label) and val and len(val) <= 40:
                rfp_meta['response_due_date'] = val
            elif 'decision' in label and 'due' in label and val and len(val) <= 40:
                rfp_meta['decision_due_date'] = val
        for row in rfp_rows[:5]:
            for cell in row:
                v = _sv(cell.value)
                if v and len(v) > 8 and 'request for proposal' not in v.lower() and 'rfp' not in v.lower()[:3]:
                    rfp_meta.setdefault('event_name', v)
                    break

    # ── Summary sheet ─────────────────────────────────────────────────────────
    if 'Summary' not in wb.sheetnames:
        return {'rfp_meta': rfp_meta, 'hotels': []}

    ws = wb['Summary']
    all_rows = list(ws.rows)

    header_row_idx = None
    for i, row in enumerate(all_rows[:10]):
        if _sv(row[0].value).lower() == 'hotel':
            header_row_idx = i
            break
    if header_row_idx is None:
        return {'rfp_meta': rfp_meta, 'hotels': []}

    headers = [_sv(all_rows[header_row_idx][i].value) for i in range(len(all_rows[header_row_idx]))]

    def _find_col(*keywords):
        for j, h in enumerate(headers):
            hl = h.lower()
            if all(kw.lower() in hl for kw in keywords):
                return j
        return None

    col_address    = _find_col('address')
    col_status     = _find_col('status')
    col_notes      = _find_col('notes')
    col_rate       = _find_col('room rate')
    col_fab        = _find_col('f&b') or _find_col('food') or _find_col('beverage') or _find_col('minimum')
    col_attrition  = _find_col('80%') or _find_col('attrition')
    col_cutoff     = _find_col('21 day') or _find_col('cut off') or _find_col('cutoff')
    col_commission = _find_col('commission') or _find_col('10%')

    def _gv(vals, idx):
        if idx is None or idx >= len(vals):
            return None
        return vals[idx]

    _SEVEN_PCT_BRANDS = [
        'marriott','westin','sheraton','w hotel','jw marriott','renaissance',
        'courtyard','residence inn','fairfield','delta hotel','tribute portfolio',
        'autograph collection','hilton','doubletree','embassy suites','hampton inn',
        'homewood','home2','curio collection','tapestry','waldorf astoria',
        'conrad','canopy','tempo by hilton','signia','graduate',
    ]
    _CONC_KEYWORDS = [
        'resort fee','comp room','1:40','1 per 40','parking','wifi',
        'wireless','internet','fitness','suite','amenity','projector','audio',
        'profit calc','f&b profit',
    ]

    skip_labels = {'available','not available','hotel',''}
    hotels = []

    for row in all_rows[header_row_idx + 1:]:
        vals = [c.value for c in row]
        hotel_name = _sv(vals[0]) if vals else ''
        if not hotel_name or hotel_name.lower() in skip_labels:
            continue
        if sum(1 for v in vals if v is not None) <= 2:
            continue

        city, state   = _parse_address(_gv(vals, col_address))
        status        = _map_status(_gv(vals, col_status))
        notes         = _sv(_gv(vals, col_notes))
        proposed_rate = _parse_currency(_gv(vals, col_rate))
        f_and_b_min   = _parse_currency(_gv(vals, col_fab))
        commission_pct = _parse_commission(_gv(vals, col_commission))

        if commission_pct is None:
            comm_raw = _sv(_gv(vals, col_commission)).lower()
            if comm_raw.startswith('yes') or '10%' in comm_raw or comm_raw == '10':
                commission_pct = 0.10
            elif any(b in hotel_name.lower() for b in _SEVEN_PCT_BRANDS):
                commission_pct = 0.07

        attrition_raw = _sv(_gv(vals, col_attrition)).lower()
        attrition_pct = 0.80 if attrition_raw.startswith('yes') else None
        cutoff_raw  = _sv(_gv(vals, col_cutoff)).lower()
        cutoff_days = 21 if cutoff_raw.startswith('yes') else None

        concession_items = []
        for j, (h, v) in enumerate(zip(headers, vals)):
            if not h or v is None:
                continue
            hl, vl = h.lower(), _sv(v).lower()
            if not any(kw in hl for kw in _CONC_KEYWORDS):
                continue
            if not vl or vl in ('no','n/a','na','—','-'):
                continue
            if vl.startswith('yes'):
                short_h = h.split('?')[0].strip()[:60]
                extra = _sv(v)[3:].strip(' ,:') if len(_sv(v)) > 3 else ''
                concession_items.append(f'✓ {short_h}: {extra}' if extra else f'✓ {short_h}')
        concessions = '\n'.join(concession_items) if concession_items else None

        crf_row_data = {h: _sv(v)[:400] for h, v in zip(headers, vals) if h and v is not None}

        hotel_dict = {
            'hotel_name': hotel_name, 'city': city, 'state': state,
            'status': status, 'proposed_rate': proposed_rate,
            'f_and_b_minimum': f_and_b_min, 'commission_pct': commission_pct,
            'attrition_pct': attrition_pct, 'cutoff_days': cutoff_days,
            'concessions': concessions, 'notes': notes or None,
            'contact_name': None, 'contact_email': None,
            'contact_phone': None, 'contact_title': None,
            'crf_row_data': _json.dumps(crf_row_data),
        }

        # Pull contact info from matching hotel tab
        h_words = set(w.lower() for w in hotel_name.split() if len(w) > 3)
        for sheet_name in wb.sheetnames:
            if sheet_name.lower() in ('summary', 'rfp'):
                continue
            s_words = set(w.lower() for w in sheet_name.split() if len(w) > 3)
            if h_words & s_words:
                try:
                    h_ws = wb[sheet_name]
                    for h_row in list(h_ws.rows)[5:22]:
                        h_vals = [c.value for c in h_row]
                        if len(h_vals) < 2:
                            continue
                        lbl = _sv(h_vals[0]).lower()
                        val = next((_sv(v) for v in h_vals[1:3] if v is not None and _sv(v)), '')
                        if not val:
                            continue
                        if 'contact name' in lbl or lbl == 'contact':
                            hotel_dict['contact_name'] = val
                        elif 'title' in lbl and 'hotel' not in lbl:
                            hotel_dict['contact_title'] = val
                        elif 'email' in lbl:
                            hotel_dict['contact_email'] = val
                        elif 'phone' in lbl or 'telephone' in lbl:
                            hotel_dict['contact_phone'] = val
                except Exception:
                    pass
                break

        hotels.append(hotel_dict)

    return {'rfp_meta': rfp_meta, 'hotels': hotels}


# ── NCSL Pickup Report Excel Importer ────────────────────────────────────────

def parse_pickup_xlsx(file_bytes):
    """
    Parse an NCSL-style pickup report Excel workbook.
    Each sheet = one event/hotel (two side-by-side grids on multi-hotel sheets).
    Returns list of dicts, one per hotel grid:
      {sheet_name, organization, hotel, event_name, contact_name, contact_email,
       booking_id, contracted_block, contracted_rate, attrition_pct, pickups}
    """
    import io as _io
    import re as _re
    from datetime import date as _date, timedelta as _td, datetime as _dt
    import openpyxl

    wb = openpyxl.load_workbook(_io.BytesIO(file_bytes), data_only=True)
    results = []

    for sheet_name in wb.sheetnames:
        if sheet_name.lower() == 'template':
            continue
        ws = wb[sheet_name]

        # Find all grid start columns: look for "PICK-UP UPDATE" in row 1
        grid_cols = []
        for col in range(1, 40):
            v = str(ws.cell(row=1, column=col).value or '').strip().upper()
            if 'PICK-UP' in v or 'PICKUP' in v:
                grid_cols.append(col)
        if not grid_cols:
            grid_cols = [1]

        for gc in grid_cols:
            grid = _parse_one_grid(ws, sheet_name, gc)
            if grid:
                results.append(grid)

    return results


def _parse_one_grid(ws, sheet_name, start_col):
    """Parse a single hotel grid starting at 1-based start_col."""
    from datetime import date as _date, timedelta as _td, datetime as _dt

    def _sv(v):
        if v is None:
            return ''
        return str(v).strip()

    def _cell(row, col_offset=0):
        return ws.cell(row=row, column=start_col + col_offset).value

    def _to_date(val):
        if val is None:
            return None
        if isinstance(val, _dt):
            return val.strftime('%Y-%m-%d')
        if hasattr(val, 'strftime'):
            return val.strftime('%Y-%m-%d')
        if isinstance(val, (int, float)) and 1000 < val < 100000:
            try:
                return (_date(1899, 12, 30) + _td(days=int(val))).strftime('%Y-%m-%d')
            except Exception:
                pass
        s = _sv(val)
        for fmt in ('%m/%d/%Y', '%Y-%m-%d', '%m/%d/%y', '%B %d, %Y'):
            try:
                from datetime import datetime as _dt2
                return _dt2.strptime(s, fmt).strftime('%Y-%m-%d')
            except Exception:
                pass
        return None

    # ── Scan label rows ───────────────────────────────────────────────────────
    org_row = hotel_row = event_row = contact_row = email_row = phone_row = None
    booking_row = dates_row = block_row = cutoff_row = None
    first_pickup_row = None

    for row in range(1, 55):
        lbl = _sv(_cell(row, 0)).lower().rstrip(':').strip()
        if 'organization' in lbl and org_row is None:
            org_row = row
        elif ('hotel' in lbl or 'location' in lbl) and hotel_row is None and 'contact' not in lbl:
            hotel_row = row
        elif ('name' in lbl and 'date' in lbl) and event_row is None:
            event_row = row
        elif 'contact' in lbl and 'number' not in lbl and 'email' not in lbl and contact_row is None:
            contact_row = row
        elif 'number' in lbl and 'booking' not in lbl and phone_row is None:
            phone_row = row
        elif 'email' in lbl and email_row is None:
            email_row = row
        elif 'booking' in lbl and booking_row is None:
            booking_row = row
        elif lbl == 'dates' and dates_row is None:
            dates_row = row
        elif lbl == 'day' and dates_row is None:
            dates_row = row
        elif 'amended block' in lbl:
            block_row = row          # prefer amended over original
        elif lbl == 'block' and block_row is None:
            block_row = row
        elif ('cut-off' in lbl or 'cutoff' in lbl) and cutoff_row is None:
            cutoff_row = row
        elif 'date called' in lbl and first_pickup_row is None:
            first_pickup_row = row

    if dates_row is None:
        return None

    # ── Header fields ─────────────────────────────────────────────────────────
    organization   = _sv(_cell(org_row, 1))    if org_row      else 'NCSL'
    hotel_name     = _sv(_cell(hotel_row, 1))  if hotel_row    else ''
    event_name     = _sv(_cell(event_row, 1))  if event_row    else ''
    contact_name   = _sv(_cell(contact_row, 1)) if contact_row else ''
    contact_phone  = _sv(_cell(phone_row, 1))  if phone_row    else ''
    contact_email  = _sv(_cell(email_row, 1))  if email_row    else ''
    booking_id     = _sv(_cell(booking_row, 1)) if booking_row else ''
    # Customer/group contact — always at col_offset 9 (label at 8) from grid start
    gc_name  = _sv(_cell(contact_row, 9)) if contact_row else ''
    gc_email = _sv(_cell(email_row, 9))   if email_row   else ''

    # ── Event dates (row 9 equivalent, cols C+ = col_offset 2+) ──────────────
    event_dates = {}   # col_offset → 'YYYY-MM-DD'
    for co in range(2, 30):
        d = _to_date(_cell(dates_row, co))
        if d:
            event_dates[co] = d
        elif event_dates:
            # Allow one None gap in case cells are merged
            d2 = _to_date(_cell(dates_row, co + 1))
            if d2 is None:
                break

    if not event_dates and dates_row is not None:
        # Some sheets have day-names on "Dates:" row and actual dates on the next "Day:" row
        for co in range(2, 30):
            d = _to_date(_cell(dates_row + 1, co))
            if d:
                event_dates[co] = d
            elif event_dates:
                d2 = _to_date(_cell(dates_row + 1, co + 1))
                if d2 is None:
                    break
        if event_dates:
            dates_row = dates_row + 1  # update so block row uses same column mapping

    if not event_dates:
        return None

    # ── Contracted block ──────────────────────────────────────────────────────
    contracted_block = {}
    if block_row:
        for co, ds in event_dates.items():
            v = _cell(block_row, co)
            if isinstance(v, (int, float)) and v > 0:
                contracted_block[ds] = int(v)

    # Remove any entry where the value = sum of all others AND > each individual value
    # (it's a "Total" column that shares a date column with a real checkout date)
    if len(contracted_block) > 2:
        block_vals = list(contracted_block.values())
        total_sum  = sum(block_vals)
        last_date  = max(contracted_block.keys())
        last_val   = contracted_block[last_date]
        rest_vals  = [v for d, v in contracted_block.items() if d != last_date]
        rest_sum   = sum(rest_vals)
        if last_val == rest_sum and last_val > max(rest_vals):
            del contracted_block[last_date]

    total_block_rooms = sum(contracted_block.values()) if contracted_block else 0

    # ── Rate & attrition from cutoff row ─────────────────────────────────────
    contracted_rate = None
    attrition_pct   = None
    attrition_rooms = None
    if cutoff_row:
        for co in range(4, 12):
            v = _cell(cutoff_row, co)
            if isinstance(v, (int, float)) and 50 < v < 2000:
                contracted_rate = float(v)
        att_v = _cell(cutoff_row, 5)
        if isinstance(att_v, (int, float)):
            attrition_pct   = float(att_v) if att_v <= 1 else att_v / 100.0
            attrition_rooms = round(total_block_rooms * attrition_pct) if total_block_rooms else None

    # ── Weekly pickup rows ────────────────────────────────────────────────────
    pickups = []
    if first_pickup_row:
        prev_total = None
        row = first_pickup_row
        max_row = first_pickup_row + 120
        while row <= max_row:
            lbl = _sv(_cell(row, 0)).lower()
            if 'date called' in lbl:
                report_date = _to_date(_cell(row, 1))
                if report_date:
                    # Pickup values are on the NEXT row
                    pr = row + 1
                    pickup_by_night = {}
                    for co, ds in event_dates.items():
                        v = _cell(pr, co)
                        if isinstance(v, (int, float)):
                            pickup_by_night[ds] = int(v)
                    # Total column = first col_offset after last date col
                    total_co = max(event_dates.keys()) + 1
                    total_v  = _cell(pr, total_co)
                    if isinstance(total_v, (int, float)):
                        total_rooms = int(total_v)
                    else:
                        total_rooms = sum(pickup_by_night.values())

                    change = (total_rooms - prev_total) if prev_total is not None else None
                    pct_block    = round(total_rooms / total_block_rooms, 4) if total_block_rooms else None
                    pct_attrition = round(total_rooms / attrition_rooms, 4) if attrition_rooms else None

                    pickups.append({
                        'report_date':    report_date,
                        'pickup_by_night': pickup_by_night,
                        'total_rooms':    total_rooms,
                        'change_from_last': change,
                        'pct_of_block':   pct_block,
                        'pct_of_attrition': pct_attrition,
                    })
                    prev_total = total_rooms
            row += 1

    return {
        'sheet_name':       sheet_name,
        'organization':     organization or 'NCSL',
        'hotel':            hotel_name,
        'event_name':       event_name,
        'contact_name':     contact_name,
        'contact_phone':    contact_phone,
        'contact_email':    contact_email,
        'gc_name':          gc_name,
        'gc_email':         gc_email,
        'booking_id':       booking_id,
        'contracted_block': contracted_block,
        'contracted_rate':  contracted_rate,
        'attrition_pct':    attrition_pct,
        'pickups':          pickups,
    }


def parse_hhr_excel(file_bytes):
    """
    Parse a Cvent Client Post Event Housing History Report (.xlsx).
    Returns a dict of summary stats.
    """
    import io as _io
    import openpyxl

    wb = openpyxl.load_workbook(_io.BytesIO(file_bytes), data_only=True)
    ws = wb.active

    def _sv(v):
        return str(v).strip() if v is not None else ''

    def _num(v):
        try:
            return float(v) if v is not None and str(v).strip() != '' else None
        except (ValueError, TypeError):
            return None

    def _to_date(v):
        if v is None:
            return None
        if hasattr(v, 'strftime'):
            return v.strftime('%Y-%m-%d')
        from datetime import date as _date, timedelta as _td
        if isinstance(v, (int, float)) and 1000 < v < 200000:
            try:
                return (_date(1899, 12, 30) + _td(days=int(v))).strftime('%Y-%m-%d')
            except Exception:
                pass
        s = _sv(v)
        for fmt in ('%Y-%m-%d', '%m/%d/%Y', '%m/%d/%y', '%B %d, %Y'):
            try:
                from datetime import datetime as _dt
                return _dt.strptime(s, fmt).strftime('%Y-%m-%d')
            except Exception:
                pass
        return None

    stats = {}
    date_cols = {}   # col_index → 'YYYY-MM-DD'
    contracted_block = {}
    final_pickup_by_night = {}

    for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
        lbl = _sv(row[0].value).lower().rstrip(':').strip() if row[0].value else ''
        col0 = _sv(row[0].value)
        col2 = row[2].value if len(row) > 2 else None

        if 'organization' in lbl and not stats.get('organization'):
            stats['organization'] = _sv(col2)
        elif 'hotel' in lbl and 'location' not in lbl and not stats.get('hotel'):
            stats['hotel'] = _sv(col2)
        elif 'name' in lbl and 'date' in lbl and not stats.get('event_name'):
            stats['event_name'] = _sv(col2)

        # Date header row — find which columns have dates
        elif 'date' == lbl and not date_cols:
            for i, cell in enumerate(row):
                if i < 2:
                    continue
                d = _to_date(cell.value)
                if d:
                    date_cols[i] = d
                elif date_cols:
                    break

        # Contracted block row
        elif 'contracted block' in lbl and not contracted_block:
            for i, d in date_cols.items():
                if i < len(row):
                    v = _num(row[i].value)
                    if v and v > 0:
                        contracted_block[d] = int(v)
            # Total in col 12 (index 12)
            if len(row) > 12:
                stats['contracted_total'] = _num(row[12].value)
            if len(row) > 13:
                stats['contracted_rate'] = _num(row[13].value)

        # Final total pickup
        elif 'final total pickup' in lbl:
            for i, d in date_cols.items():
                if i < len(row):
                    v = _num(row[i].value)
                    if v and v > 0:
                        final_pickup_by_night[d] = int(v)
            if len(row) > 12:
                stats['final_total_pickup'] = _num(row[12].value)

        # Total pickup inside block
        elif 'total pickup inside block' in lbl:
            if len(row) > 12:
                stats['pickup_inside_block'] = _num(row[12].value)

        # Audit pickup
        elif 'total audit pickup' in lbl:
            if len(row) > 12:
                stats['audit_pickup'] = _num(row[12].value)

        # No shows / Cancellations
        elif 'no show' in lbl:
            if len(row) > 12:
                stats['no_shows'] = _num(row[12].value)
        elif 'cancellation' in lbl:
            if len(row) > 12:
                stats['cancellations'] = _num(row[12].value)

        # Revenue lines — labels may appear in any column, scan the whole row
        row_text = ' '.join(_sv(c.value).lower() for c in row)
        if 'total actualized pickup revenue' in row_text and not stats.get('room_revenue'):
            for i, cell in enumerate(row):
                if 'total actualized pickup revenue' in _sv(cell.value).lower():
                    # value is 2 columns to the right
                    for j in (i+2, i+1, i+3):
                        if j < len(row):
                            v = _num(row[j].value)
                            if v and v > 0:
                                stats['room_revenue'] = v
                                break
                    break
        if ('total food' in row_text or ('food' in row_text and 'beverage' in row_text)) and not stats.get('fb_revenue'):
            for i, cell in enumerate(row):
                if 'total food' in _sv(cell.value).lower() or ('food' in _sv(cell.value).lower() and 'beverage' in _sv(cell.value).lower()):
                    for j in (i+2, i+1, i+3):
                        if j < len(row):
                            v = _num(row[j].value)
                            if v and v > 0:
                                stats['fb_revenue'] = v
                                break
                    break

        # Earned comps
        elif 'total comp amount' in lbl or ('earned comp' in lbl and 'total' in lbl):
            if len(row) > 12:
                stats['earned_comps_value'] = _num(row[12].value)
        elif 'total rns eligible for earned comp' in lbl:
            if len(row) > 12:
                stats['earned_comps_rns'] = _num(row[12].value)

        # Hotel accounting sign-off
        elif 'history report approved' in lbl:
            for i, cell in enumerate(row):
                v = _sv(cell.value)
                if v.lower().startswith('matthew') or (i > 0 and _sv(row[i-1].value).lower() == 'name'):
                    pass
            # Scan for Email field
            for i, cell in enumerate(row):
                if _sv(cell.value).lower() == 'email' and i+1 < len(row):
                    stats['hotel_approver_email'] = _sv(row[i+1].value)
                if _sv(cell.value).lower() == 'name' and i+1 < len(row):
                    stats['hotel_approver'] = _sv(row[i+1].value)
                d = _to_date(cell.value)
                if d and not stats.get('report_date'):
                    stats['report_date'] = d

    stats['contracted_block']     = contracted_block
    stats['final_pickup_by_night'] = final_pickup_by_night

    # Derived stats
    ct = stats.get('contracted_total') or sum(contracted_block.values()) or None
    fp = stats.get('final_total_pickup')
    if ct and fp:
        stats['pct_of_block'] = round(fp / ct * 100, 1)
    if ct and fp and stats.get('contracted_rate'):
        stats['room_revenue_calc'] = round(fp * stats['contracted_rate'], 2)

    return stats


def strip_hhr_commission_rows(file_bytes):
    """
    Return a client-safe copy of the HHR Excel file with commission data removed.

    Commission labels live in column O (15); values in column Q (17).
    Actions:
      • Rename "Total Actualized Pickup Revenue (Inside Block)" → "Total Actualized Revenue"
        and update Q of that row with the total actualized revenue
        (inside block + audit pickup, i.e. the gross commissionable revenue).
      • Clear O + Q for all other commission rows.
      • Rename section header "ROOM REVENUE/COMMISSION" → "ROOM REVENUE".
      • Erase the "Commission Check Generated By" row (col A) entirely.
    """
    import io as _io
    from openpyxl import load_workbook

    LABEL_COL = 15   # O — commission label column
    VALUE_COL = 17   # Q — commission value column

    LABELS_TO_CLEAR = {
        'total actualized pickup commission (inside block)',
        'total commissionable audit pickup revenue',
        'total commissionable audit pickup commission',
        'total commissionable no show commission',
        'total commissionable cancellation commission',
        'less housing fee commission',
        'less rebate commission',
        'less earned comp commission',
        'total rooms commission due',
    }

    # Load with data_only so formula results are available as values.
    # (Saving will convert all formulas to static values — acceptable for client delivery.)
    wb = load_workbook(_io.BytesIO(file_bytes), data_only=True)
    ws = wb.active

    # ── First pass: collect values we need + find AUDIT DETAIL row ───────────
    inside_revenue   = 0.0
    audit_revenue    = 0.0
    audit_detail_row = None   # first row of the AUDIT DETAIL section

    for row in ws.iter_rows():
        rnum  = row[0].row
        o_val = str(ws.cell(row=rnum, column=LABEL_COL).value or '').strip()
        o_lbl = o_val.lower()
        q_val = ws.cell(row=rnum, column=VALUE_COL).value

        # Detect AUDIT DETAIL header (appears in column D / col 4)
        d_val = str(ws.cell(row=rnum, column=4).value or '').strip().lower()
        if d_val == 'audit detail' and audit_detail_row is None:
            audit_detail_row = rnum

        if o_lbl == 'total actualized pickup revenue (inside block)':
            inside_revenue = float(q_val or 0)
        elif o_lbl == 'total commissionable audit pickup revenue':
            audit_revenue  = float(q_val or 0)

    gross_revenue = inside_revenue + audit_revenue

    # ── Delete AUDIT DETAIL section (all rows from that header to end) ────────
    if audit_detail_row:
        ws.delete_rows(audit_detail_row, ws.max_row - audit_detail_row + 1)

    # ── Second pass: apply commission changes ──────────────────────────────────
    for row in ws.iter_rows():
        rnum   = row[0].row
        o_cell = ws.cell(row=rnum, column=LABEL_COL)
        q_cell = ws.cell(row=rnum, column=VALUE_COL)
        a_cell = ws.cell(row=rnum, column=1)
        o_lbl  = str(o_cell.value or '').strip().lower()
        a_lbl  = str(a_cell.value or '').strip().lower()

        if o_lbl == 'room revenue/commission':
            o_cell.value = 'ROOM REVENUE'

        elif o_lbl == 'total actualized pickup revenue (inside block)':
            o_cell.value = 'Total Actualized Revenue'
            q_cell.value = gross_revenue if gross_revenue else inside_revenue

        elif o_lbl in LABELS_TO_CLEAR:
            o_cell.value = None
            q_cell.value = None

        if 'commission check generated' in a_lbl:
            # Unmerge any merged cells in this row before clearing
            ranges_to_remove = [mr for mr in ws.merged_cells.ranges if mr.min_row <= rnum <= mr.max_row]
            for mr in ranges_to_remove:
                ws.unmerge_cells(str(mr))
            for col in range(1, ws.max_column + 1):
                cell = ws.cell(row=rnum, column=col)
                try:
                    cell.value = None
                except AttributeError:
                    pass

    out = _io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.read()
