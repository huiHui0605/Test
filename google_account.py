"""Desktop OAuth and narrowly scoped Google API access. Never log API payloads."""
import json
import os
import threading
from pathlib import Path

SCOPES = (
    'openid', 'https://www.googleapis.com/auth/userinfo.email',
    'https://www.googleapis.com/auth/drive.metadata.readonly',
    'https://www.googleapis.com/auth/spreadsheets.readonly',
    'https://www.googleapis.com/auth/gmail.send',
)
SERVICE = 'Foodbank Community Admin Google OAuth'


class GoogleError(RuntimeError):
    """Safe, operator-facing error without response bodies or credentials."""


def safe_error(error):
    from google.auth.exceptions import RefreshError
    if isinstance(error, GoogleError):
        return str(error)
    if type(error).__name__ == 'WSGITimeoutError':
        return 'Google sign-in timed out. Click Connect Google and complete the browser authorization within three minutes.'
    if type(error).__name__ == 'AccessDeniedError':
        return 'Google sign-in was cancelled or access was declined. Click Connect Google when ready.'
    status = getattr(getattr(error, 'resp', None), 'status', None)
    if isinstance(error, RefreshError) or status == 401:
        return 'Google authorization expired or was revoked. Click Connect Google to sign in again.'
    if status == 403:
        return 'Google denied access. Check granted permissions, API setup, and access to this spreadsheet or Gmail mailbox.'
    if status == 404:
        return 'The selected Google resource is unavailable. Check account access and select the source again.'
    return 'Google could not complete this request. Check your connection and Google setup, then try again.'


class CredentialVault:
    """Only OS-backed stores are accepted; never fall back to a plaintext file."""
    def __init__(self, slot, backend=None):
        self.slot = slot
        if backend is None:
            import keyring
            if os.name == 'nt':
                from keyring.backends.Windows import WinVaultKeyring
                backend = WinVaultKeyring()
            elif __import__('sys').platform == 'darwin':
                from keyring.backends.macOS import Keyring
                backend = Keyring()
            else:
                backend = keyring.get_keyring()
                if type(backend).__module__ not in ('keyring.backends.SecretService', 'keyring.backends.kwallet'):
                    raise GoogleError('No supported secure credential store is available. Enable Windows Credential Manager, macOS Keychain, or Linux Secret Service/KWallet.')
        self.backend = backend

    def read(self):
        try:
            return self.backend.get_password(SERVICE, self.slot)
        except Exception:
            raise GoogleError('Cannot unlock the operating-system credential store. Unlock it and try again.') from None

    def save(self, value):
        try:
            self.backend.set_password(SERVICE, self.slot, value)
        except Exception:
            raise GoogleError('Cannot save Google authorization securely. Check the operating-system credential store.') from None

    def clear(self):
        try:
            if self.read() is not None:
                self.backend.delete_password(SERVICE, self.slot)
        except GoogleError:
            raise
        except Exception:
            raise GoogleError('Cannot remove Google authorization from the operating-system credential store.') from None


def client_configuration(path):
    try:
        config = json.loads(Path(path).read_text(encoding='utf-8'))
        client = config['installed']
        if not client.get('client_id') or not client.get('client_secret'):
            raise ValueError()
        # Do not send credentials to endpoints supplied by an arbitrary JSON file.
        if client.get('auth_uri') != 'https://accounts.google.com/o/oauth2/auth' or client.get('token_uri') != 'https://oauth2.googleapis.com/token':
            raise ValueError()
        return config
    except Exception:
        raise GoogleError('Ask the technical administrator to configure a Google OAuth Desktop App client JSON in Developer Google Setup. Service-account keys are not supported.') from None


class SingleAttemptHttp:
    """Google API transport for sends: no retry, redirect, or auth replay.

    httplib2 retries a stale connection's BadStatusLine even when its retry
    setting is reduced. A fresh requests session with retries disabled avoids
    silently replaying an already submitted Gmail POST.
    """
    def __init__(self, credentials):
        import requests
        from requests.adapters import HTTPAdapter
        self.credentials = credentials
        self.session = requests.Session()
        self.session.mount('https://', HTTPAdapter(max_retries=0))

    def request(self, uri, method='GET', body=None, headers=None, **kwargs):
        import httplib2
        from urllib.parse import urlsplit
        parsed = urlsplit(uri)
        if parsed.scheme != 'https' or parsed.hostname not in ('gmail.googleapis.com', 'www.googleapis.com'):
            raise GoogleError('Unexpected Gmail API endpoint. No email was sent.')
        headers = dict(headers or {})
        self.credentials.apply(headers)
        result = self.session.request(method, uri, data=body, headers=headers, timeout=30, allow_redirects=False)
        response = httplib2.Response(dict(result.headers, status=str(result.status_code)))
        return response, result.content

    def close(self):
        self.session.close()


class GoogleAccount:
    def __init__(self, slot, client_path='', vault=None):
        self.slot = slot
        self.client_path = client_path or os.environ.get('FOODBANK_GOOGLE_CLIENT', '')
        self._vault = vault
        self._lock = threading.RLock()

    @property
    def vault(self):
        if self._vault is None:
            self._vault = CredentialVault(self.slot)
        return self._vault

    def credentials(self):
        from google.oauth2.credentials import Credentials
        from google.auth.transport.requests import Request
        with self._lock:
            saved = self.vault.read()
            if not saved:
                raise GoogleError('Google is disconnected. Click Connect Google to sign in.')
            try:
                info = json.loads(saved)
                if info.get('token_uri') != 'https://oauth2.googleapis.com/token':
                    raise ValueError()
                cred = Credentials.from_authorized_user_info(info)
                if not cred.has_scopes(SCOPES):
                    raise GoogleError('Google permissions have changed. Click Connect Google and grant all requested permissions.')
                if not cred.valid:
                    if not cred.refresh_token:
                        raise GoogleError('Click Connect Google to sign in again.')
                    request = Request()
                    cred.refresh(lambda *a, **kw: request(*a, **dict(kw, timeout=30)))
                    self.vault.save(cred.to_json())
                return cred
            except Exception as error:
                raise GoogleError(safe_error(error)) from None

    def connect(self):
        from google_auth_oauthlib.flow import InstalledAppFlow
        # Check the vault before opening a browser; do not offer insecure storage.
        self.vault.read()
        config = client_configuration(self.client_path)
        try:
            flow = InstalledAppFlow.from_client_config(config, SCOPES, autogenerate_code_verifier=True)
            request = flow.oauth2session.request
            flow.oauth2session.request = lambda *a, **kw: request(*a, **dict(kw, timeout=30))
            cred = flow.run_local_server(
                host='127.0.0.1', port=0, open_browser=True, timeout_seconds=180,
                authorization_prompt_message='',
                success_message='Google authorization received. You can close this window and return to Foodbank.',
                access_type='offline', prompt='consent select_account',
            )
            if not cred.refresh_token or not cred.has_scopes(SCOPES):
                raise GoogleError('Sign-in did not grant all required permissions. Click Connect Google and allow the requested access.')
            # Check the actual granted scopes when Google's token response supplies them.
            granted = cred.granted_scopes
            if granted is not None and not set(SCOPES).issubset(set(granted)):
                raise GoogleError('Some Google permissions were declined. Reconnect and grant all requested permissions.')
            identity = self.identity(credentials=cred)
            with self._lock:
                self.vault.save(cred.to_json())
            return identity
        except Exception as error:
            raise GoogleError(safe_error(error)) from None

    def disconnect(self):
        # Local disconnect is deliberately offline. Google-account revocation is documented.
        with self._lock:
            self.vault.clear()

    def service(self, name, version, credentials=None):
        from googleapiclient.discovery import build
        from google_auth_httplib2 import AuthorizedHttp
        import httplib2
        cred = credentials or self.credentials()
        if name == 'gmail':
            return build(name, version, http=SingleAttemptHttp(cred), cache_discovery=False)
        http = AuthorizedHttp(cred, http=httplib2.Http(timeout=30, ca_certs=os.environ.get('SSL_CERT_FILE') or os.environ.get('REQUESTS_CA_BUNDLE')), max_refresh_attempts=0)
        return build(name, version, http=http, cache_discovery=False)

    def identity(self, credentials=None):
        try:
            result = self.service('oauth2', 'v2', credentials=credentials).userinfo().get().execute(num_retries=0)
            return result.get('email', 'Google account')
        except Exception as error:
            raise GoogleError(safe_error(error)) from None

    def spreadsheets(self):
        try:
            service = self.service('drive', 'v3')
            results = []; token = None
            while True:
                page = service.files().list(
                    q="mimeType='application/vnd.google-apps.spreadsheet' and trashed=false",
                    spaces='drive', fields='nextPageToken,files(id,name)', pageSize=100,
                    pageToken=token, supportsAllDrives=True, includeItemsFromAllDrives=True,
                ).execute(num_retries=0)
                results.extend(page.get('files', [])); token = page.get('nextPageToken')
                if not token:
                    return sorted(results, key=lambda x: (x['name'].casefold(), x['id']))
        except Exception as error:
            raise GoogleError(safe_error(error)) from None

    def worksheets(self, spreadsheet):
        try:
            result = self.service('sheets', 'v4').spreadsheets().get(
                spreadsheetId=spreadsheet, fields='sheets(properties(sheetId,title,gridProperties(columnCount)))',
            ).execute(num_retries=0)
            return [s['properties'] for s in result.get('sheets', [])]
        except Exception as error:
            raise GoogleError(safe_error(error)) from None

    def headers(self, spreadsheet, worksheet):
        from google_sync import worksheet_range
        try:
            result = self.service('sheets', 'v4').spreadsheets().values().get(
                spreadsheetId=spreadsheet, range=worksheet_range(worksheet, header_only=True), valueRenderOption='FORMATTED_VALUE',
            ).execute(num_retries=0)
            grid = result.get('values', [])
            headers = [str(h).strip() for h in grid[0]] if grid else []
            if not headers or any(not h or h.startswith('__fb_') for h in headers) or len(set(headers)) != len(headers):
                raise GoogleError('The first worksheet row must contain nonempty, unique question headings, without reserved __fb_ names.')
            return headers
        except Exception as error:
            raise GoogleError(safe_error(error)) from None
