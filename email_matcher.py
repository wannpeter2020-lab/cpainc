"""
email_matcher.py — shared helper for matching emails to active pickup/RFP events.

Provides:
  get_event_terms(db_path)  → list of term dicts
  match_email(subject, sender, preview, terms) → matched term dict or None
"""

import re
import sqlite3

DB_PATH_DEFAULT = None  # set by caller or auto-detected

# Common words that should NOT be used as standalone match keywords
# (too generic — would cause false positives)
_STOPWORDS = {
    # Hotel brand / property words
    'hotel', 'hotels', 'suite', 'suites', 'resort', 'resorts', 'inn', 'inns',
    'marriott', 'hilton', 'hyatt', 'westin', 'sheraton', 'embassy', 'autograph',
    'collection', 'center', 'centre', 'grand', 'regency', 'courtyard', 'plaza',
    'downtown', 'airport', 'extended', 'stay', 'select', 'garden', 'doubletree',
    'crowne', 'western', 'express', 'springhill', 'homewood', 'hampton', 'comfort',
    'quality', 'sleep', 'wingate', 'wyndham', 'radisson', 'intercontinental',
    'kimpton', 'omni', 'loews', 'dream', 'origin', 'indigo', 'voco', 'curio',
    'tapestry', 'tribute', 'boutique', 'luxury',
    # Common English words
    'the', 'and', 'for', 'with', 'from', 'that', 'this', 'which', 'into',
    'north', 'south', 'east', 'west', 'lake', 'creek', 'house', 'manor',
    'club', 'park', 'city', 'museum', 'lounge', 'country', 'best', 'four',
    'seasons', 'national', 'international', 'american', 'annual',
    # Event / business words too generic to match on
    'group', 'event', 'events', 'meeting', 'conference', 'leadership', 'excellence',
    'follow', 'sourcing', 'update', 'proposal', 'rfp', 'booking', 'spring', 'summer',
    # Words common across multiple FCCS events — too noisy as standalone terms
    'leading', 'leaders', 'large', 'other', 'others', 'role', 'board', 'boards',
    'strategy', 'learning', 'forum', 'self', 'system', 'director', 'sales',
    # Person first names that appear in org fields — match by full name only
    'matt', 'lori', 'peter', 'john', 'mike', 'chris', 'david', 'james', 'robert',
}


# ── Marketing / newsletter detection ─────────────────────────────────────────

# Text patterns found in the body preview of marketing/newsletter emails.
# Presence of any of these means "skip this email."
_MARKETING_PREVIEW_PATTERNS = re.compile(
    r'unsubscribe|opt[- ]out|email preferences|manage (your )?preferences|'
    r'view (in|this email) (browser|online)|view online|'
    r'you(?:\'re| are) receiving this|why (am i|did i) (get|receive)|'
    r'update (your )?email|privacy policy|terms of (use|service)|'
    r'you have received this|sent to .{0,60} because|'
    r'this is a (promotional|marketing|commercial) email|'
    r'no longer wish to receive|click here to (unsubscribe|manage)|'
    r'to stop receiving|remove (me|yourself) from|mailing list',
    re.IGNORECASE,
)

# Sender address domains that are exclusively bulk/marketing platforms.
_MARKETING_SENDER_DOMAINS = {
    'mailchimp.com', 'list-manage.com', 'constantcontact.com',
    'sendgrid.net', 'sendgrid.com', 'klaviyo.com', 'klaviyomail.com',
    'hubspot.com', 'hs-mail.com', 'marketo.com', 'pardot.com',
    'eloqua.com', 'exacttarget.com', 'salesforceiq.com',
    'sailthru.com', 'sailthru.net', 'yesmail.com', 'bluehornet.com',
    'aweber.com', 'getresponse.com', 'campaignmonitor.com',
    'createsend.com', 'dotmailer.com', 'dotdigital.com',
    'mailerlite.com', 'drip.com', 'activecampaign.com',
    'infusionsoft.com', 'keap.com', 'ontraport.com',
    'emma.email', 'myemma.com', 'benchmarkemail.com',
    'icontact.com', 'verticalresponse.com', 'freshmail.com',
    'omnisend.com', 'sendinblue.com', 'brevo.com',
    'postmarkapp.com', 'sparkpostmail.com', 'amazonses.com',
    'convio.net', 'blackbaudhq.com',
}


def is_marketing_email(sender_address: str, preview: str) -> bool:
    """
    Return True if this email looks like a newsletter or marketing message.
    Checks:
      1. Sender domain matches a known bulk-mail platform.
      2. Body preview contains standard unsubscribe / marketing language.
    """
    # Check sender domain
    addr = (sender_address or '').lower()
    domain = addr.split('@')[-1].strip('>')
    if domain in _MARKETING_SENDER_DOMAINS:
        return True

    # Check preview text
    if preview and _MARKETING_PREVIEW_PATTERNS.search(preview):
        return True

    return False


def _extract_keywords(term):
    """
    Extract standalone match keywords from a term string.
    - Parenthetical shorthands like (LLO 15-4) or (LO 42-2) extracted as compound terms
    - All-caps acronyms >= 3 chars (e.g. AGS, BRS, ECS, LLO)
    - Any word >= 5 chars not in stopwords
    Returns a list of unique keywords/phrases (preserving original case).
    """
    result = []
    seen = set()

    # 1. Extract parenthetical shorthands first — e.g. (LLO 15-4), (LO 42-2)
    #    These are the most specific identifiers and should rank highest.
    for paren in re.findall(r'\(([^)]+)\)', term):
        paren = paren.strip()
        if len(paren) >= 3 and paren.lower() not in seen:
            seen.add(paren.lower())
            result.append(paren)

    # 2. Individual words
    words = re.findall(r"[A-Za-z]+", term)
    for w in words:
        wl = w.lower()
        if wl in seen or wl in _STOPWORDS:
            continue
        is_acronym = w.isupper() and len(w) >= 3
        is_long_word = len(w) >= 5
        if (is_acronym or is_long_word) and wl != term.lower():
            seen.add(wl)
            result.append(w)
    return result


def get_db(db_path):
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    return conn


def get_event_terms(db_path):
    """
    Return a list of dicts describing active events to match against:
      { term, label, event_type, event_id, url }

    Sources:
      - pickup_config: active pickup events (organization, event_name, hotel)
      - rfp: active RFPs (client_org, event_name)
      - rfp_hotel: hotel names on active RFPs
      - ReportPipeline: active bookings (ClientName, EventName, HotelName)
    """
    terms = []
    conn = get_db(db_path)

    def _add_terms(term_list, val, label, event_type, event_id, url, color,
                   extract_keywords=False):
        """Append the full term + optional individual keywords to term_list."""
        val = (val or '').strip()
        if not val:
            return
        base = {
            'label': label, 'event_type': event_type,
            'event_id': event_id, 'url': url, 'color': color,
        }
        # Full term (min 4 chars, or 3-char all-caps acronym)
        if len(val) >= 4 or (val.isupper() and len(val) >= 3):
            term_list.append({**base, 'term': val})
        # Individual keywords extracted from multi-word terms
        if extract_keywords and ' ' in val:
            for kw in _extract_keywords(val):
                term_list.append({**base, 'term': kw})

    # ── Pickup events ──────────────────────────────────────────────────────────
    try:
        rows = conn.execute(
            "SELECT id, organization, event_name, hotel FROM pickup_config "
            "WHERE status IS NULL OR status NOT IN ('cancelled','lost','closed')"
        ).fetchall()
        for r in rows:
            eid = r['id']
            label_base = r['event_name'] or r['organization']
            url = f'/pickup/{eid}'
            for col in ('event_name', 'organization'):
                _add_terms(terms, r[col], label_base, 'pickup', eid, url, 'primary',
                           extract_keywords=True)
            # Hotel: full name + keyword extraction
            _add_terms(terms, r['hotel'], label_base, 'pickup', eid, url, 'primary',
                       extract_keywords=True)
    except Exception:
        pass

    # ── RFP events ────────────────────────────────────────────────────────────
    try:
        rows = conn.execute(
            "SELECT id, client_org, event_name FROM rfp "
            "WHERE status NOT IN ('Cancelled','Lost','Closed') OR status IS NULL"
        ).fetchall()
        for r in rows:
            eid = r['id']
            label_base = r['event_name'] or r['client_org']
            url = f'/rfp/{eid}'
            for col in ('event_name', 'client_org'):
                _add_terms(terms, r[col], label_base, 'rfp', eid, url, 'warning',
                           extract_keywords=True)
    except Exception:
        pass

    # ── RFP hotel names ───────────────────────────────────────────────────────
    try:
        rows = conn.execute(
            "SELECT rh.rfp_id, rh.hotel_name, r.event_name, r.client_org "
            "FROM rfp_hotel rh JOIN rfp r ON r.id = rh.rfp_id "
            "WHERE r.status NOT IN ('Cancelled','Lost','Closed') OR r.status IS NULL"
        ).fetchall()
        for r in rows:
            label = r['event_name'] or r['client_org']
            url   = f'/rfp/{r["rfp_id"]}'
            # Full hotel name + unique keywords (e.g. "Jacquard", "Halcyon")
            _add_terms(terms, r['hotel_name'], label, 'rfp', r['rfp_id'], url, 'warning',
                       extract_keywords=True)
    except Exception:
        pass

    # ── Pipeline bookings ─────────────────────────────────────────────────────
    try:
        rows = conn.execute(
            "SELECT id, ClientName, EventName, HotelName FROM ReportPipeline "
            "WHERE Status NOT IN ('Cancelled','Lost') OR Status IS NULL"
        ).fetchall()
        for r in rows:
            eid = r['id']
            label_base = r['EventName'] or r['ClientName']
            url = f'/booking/{eid}'
            for col in ('EventName', 'ClientName'):
                _add_terms(terms, r[col], label_base, 'booking', eid, url, 'info',
                           extract_keywords=True)
            _add_terms(terms, r['HotelName'], label_base, 'booking', eid, url, 'info',
                       extract_keywords=True)
    except Exception:
        pass

    conn.close()

    # Deduplicate by (term.lower(), event_id, event_type)
    seen = set()
    unique = []
    for t in terms:
        key = (t['term'].lower(), t['event_id'], t['event_type'])
        if key not in seen:
            seen.add(key)
            unique.append(t)

    return unique


def match_email(subject, sender, preview, terms):
    """
    Check if an email matches any event term.
    Returns the best matching term dict, or None.
    Prioritises longer/more-specific terms.
    """
    text = f'{subject} {sender} {preview}'.lower()
    matches = []
    for t in terms:
        word = t['term'].lower()
        # Require at least a word-boundary-ish match (not just substring of another word)
        if re.search(r'(?<![a-z])' + re.escape(word) + r'(?![a-z])', text):
            matches.append(t)

    if not matches:
        return None

    # Return the most specific match:
    # Primary key: length of term (longer = more specific)
    # Tiebreaker: prefer event_name/org matches over hotel matches (event_type wins ties)
    return max(matches, key=lambda t: (len(t['term']), t['event_type'] != 'booking'))


def match_email_all(subject, sender, preview, terms):
    """Return all matching term dicts (deduplicated by event_id+type)."""
    text = f'{subject} {sender} {preview}'.lower()
    seen = set()
    results = []
    for t in terms:
        word = t['term'].lower()
        if re.search(r'(?<![a-z])' + re.escape(word) + r'(?![a-z])', text):
            key = (t['event_id'], t['event_type'])
            if key not in seen:
                seen.add(key)
                results.append(t)
    return results
