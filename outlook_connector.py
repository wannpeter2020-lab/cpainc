"""
Microsoft Graph API connector for Promagent.
Handles OAuth2 auth, email, calendar, and contacts via Microsoft 365.
"""

import json
import os
import re
import requests
from datetime import datetime, timedelta, timezone

try:
    import msal
except ImportError:
    msal = None

# ── Config ────────────────────────────────────────────────────────────────────
_here = os.path.dirname(__file__)
TOKEN_FILE = os.path.join(_here, 'outlook_token.json')

GRAPH_BASE = 'https://graph.microsoft.com/v1.0'

SCOPES = [
    'Mail.ReadWrite',
    'Mail.Send',
    'Calendars.ReadWrite',
    'Contacts.ReadWrite',
    'User.Read',
]


def _load_config():
    try:
        from config import MS_CLIENT_ID, MS_TENANT_ID, MS_CLIENT_SECRET
        return MS_CLIENT_ID, MS_TENANT_ID, MS_CLIENT_SECRET
    except ImportError:
        return None, None, None


def _get_app():
    client_id, tenant_id, client_secret = _load_config()
    if not client_id or client_id == 'YOUR_CLIENT_ID':
        return None
    if msal is None:
        return None
    authority = f'https://login.microsoftonline.com/{tenant_id}'
    return msal.ConfidentialClientApplication(
        client_id,
        authority=authority,
        client_credential=client_secret,
    )


# ── Auth ──────────────────────────────────────────────────────────────────────

def get_auth_url(redirect_uri):
    app = _get_app()
    if app is None:
        return None
    return app.get_authorization_request_url(SCOPES, redirect_uri=redirect_uri)


def exchange_code_for_token(code, redirect_uri):
    app = _get_app()
    if app is None:
        return False, 'Microsoft credentials not configured'
    result = app.acquire_token_by_authorization_code(code, scopes=SCOPES, redirect_uri=redirect_uri)
    if 'access_token' in result:
        _save_token(result)
        return True, None
    return False, result.get('error_description', 'Unknown error')


def _save_token(token_data):
    with open(TOKEN_FILE, 'w') as f:
        json.dump(token_data, f)


def _load_token():
    if not os.path.exists(TOKEN_FILE):
        return None
    with open(TOKEN_FILE) as f:
        return json.load(f)


def get_token():
    token = _load_token()
    if not token:
        return None

    # Try refresh_token grant
    refresh_token = token.get('refresh_token')
    if refresh_token:
        client_id, tenant_id, client_secret = _load_config()
        url = f'https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token'
        data = {
            'grant_type': 'refresh_token',
            'client_id': client_id,
            'client_secret': client_secret,
            'refresh_token': refresh_token,
            'scope': ' '.join(SCOPES),
        }
        resp = requests.post(url, data=data)
        if resp.status_code == 200:
            new_token = resp.json()
            if 'refresh_token' not in new_token:
                new_token['refresh_token'] = refresh_token
            _save_token(new_token)
            return new_token['access_token']

    return token.get('access_token')


def is_connected():
    return os.path.exists(TOKEN_FILE)


def disconnect():
    if os.path.exists(TOKEN_FILE):
        os.remove(TOKEN_FILE)


def _headers():
    token = get_token()
    if not token:
        raise RuntimeError('Not authenticated with Microsoft. Visit /outlook/auth to connect.')
    return {'Authorization': f'Bearer {token}', 'Content-Type': 'application/json'}


def _graph(method, path, **kwargs):
    url = GRAPH_BASE + path
    resp = requests.request(method, url, headers=_headers(), **kwargs)
    resp.raise_for_status()
    return resp.json() if resp.content else {}


# ── User info ─────────────────────────────────────────────────────────────────

def get_user_info():
    try:
        data = _graph('GET', '/me', params={'$select': 'displayName,mail,userPrincipalName'})
        return {
            'name': data.get('displayName', ''),
            'email': data.get('mail') or data.get('userPrincipalName', ''),
        }
    except Exception:
        return None


# ── Email ─────────────────────────────────────────────────────────────────────

def send_email(to, subject, body_html, cc=None):
    to_list = [to] if isinstance(to, str) else list(to)
    message = {
        'subject': subject,
        'body': {'contentType': 'HTML', 'content': body_html},
        'toRecipients': [{'emailAddress': {'address': a}} for a in to_list],
    }
    if cc:
        cc_list = [cc] if isinstance(cc, str) else list(cc)
        message['ccRecipients'] = [{'emailAddress': {'address': a}} for a in cc_list]
    _graph('POST', '/me/sendMail', json={'message': message, 'saveToSentItems': True})


def get_emails(folder='Inbox', count=20, filter_query=None, search=None):
    params = {
        '$top': count,
        '$select': 'id,subject,from,receivedDateTime,bodyPreview,isRead,webLink',
        '$orderby': 'receivedDateTime desc',
    }
    if filter_query:
        params['$filter'] = filter_query
    if search:
        params['$search'] = f'"{search}"'
        params.pop('$orderby', None)
    data = _graph('GET', f'/me/mailFolders/{folder}/messages', params=params)
    return data.get('value', [])


def mark_as_read(message_id):
    try:
        requests.patch(f'{GRAPH_BASE}/me/messages/{message_id}',
                       headers=_headers(), json={'isRead': True})
    except Exception:
        pass


# ── Calendar ──────────────────────────────────────────────────────────────────

def get_calendar_events(days_ahead=30, days_back=0):
    now = datetime.now(timezone.utc)
    start = (now - timedelta(days=days_back)).strftime('%Y-%m-%dT%H:%M:%SZ')
    end   = (now + timedelta(days=days_ahead)).strftime('%Y-%m-%dT%H:%M:%SZ')
    params = {
        'startDateTime': start,
        'endDateTime': end,
        '$select': 'id,subject,start,end,location,bodyPreview,webLink',
        '$orderby': 'start/dateTime',
        '$top': 50,
    }
    token = get_token()
    if not token:
        raise RuntimeError('Not authenticated with Microsoft.')
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json',
        'Prefer': 'outlook.timezone="Eastern Standard Time"',
    }
    resp = requests.get(f'{GRAPH_BASE}/me/calendarView', headers=headers, params=params)
    resp.raise_for_status()
    return resp.json().get('value', [])


def create_calendar_event(subject, start_dt, end_dt, location='', body='', all_day=False):
    def _fmt(dt):
        if isinstance(dt, datetime):
            return dt.strftime('%Y-%m-%dT%H:%M:%S')
        return str(dt)
    event = {
        'subject': subject,
        'start': {'dateTime': _fmt(start_dt), 'timeZone': 'Pacific Standard Time'},
        'end':   {'dateTime': _fmt(end_dt),   'timeZone': 'Pacific Standard Time'},
        'body':  {'contentType': 'HTML', 'content': body},
        'isAllDay': all_day,
    }
    if location:
        event['location'] = {'displayName': location}
    return _graph('POST', '/me/events', json=event)


# ── Contacts ──────────────────────────────────────────────────────────────────

def get_contacts(count=100):
    params = {
        '$top': count,
        '$select': 'id,displayName,emailAddresses,companyName,businessPhones,jobTitle',
        '$orderby': 'displayName',
    }
    data = _graph('GET', '/me/contacts', params=params)
    return data.get('value', [])


def sync_contact(display_name, email, company='', phone='', job_title=''):
    contacts = get_contacts()
    existing = next(
        (c for c in contacts if any(e.get('address', '').lower() == email.lower()
                                    for e in c.get('emailAddresses', []))),
        None
    )
    payload = {
        'displayName': display_name,
        'companyName': company,
        'jobTitle': job_title,
        'emailAddresses': [{'address': email, 'name': display_name}],
    }
    if phone:
        payload['businessPhones'] = [phone]
    if existing:
        requests.patch(f'{GRAPH_BASE}/me/contacts/{existing["id"]}',
                       headers=_headers(), json=payload)
        return 'updated'
    else:
        _graph('POST', '/me/contacts', json=payload)
        return 'created'
