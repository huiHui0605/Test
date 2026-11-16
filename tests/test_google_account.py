import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch
from datetime import datetime, timedelta, timezone
from google_account import GoogleAccount, CredentialVault, GoogleError, SCOPES, SERVICE, client_configuration
from google_sync import fetch, validate_source, worksheet_range
from storage import ExcelStore


class MemoryVault:
    def __init__(self, value=None):self.value=value
    def read(self):return self.value
    def save(self,value):self.value=value
    def clear(self):self.value=None


def token_info(expired=False):
    expiry=datetime.now(timezone.utc)+timedelta(hours=-1 if expired else 1)
    return json.dumps(dict(token='fictional-access-token',refresh_token='fictional-refresh-token',
        client_id='fictional-client',client_secret='fictional-secret',token_uri='https://oauth2.googleapis.com/token',
        scopes=list(SCOPES),expiry=expiry.strftime('%Y-%m-%dT%H:%M:%SZ')))


class OAuthTests(unittest.TestCase):
    def test_secure_os_vault_roundtrip_and_disconnect(self):
        backend=Mock();backend.get_password.return_value=token_info()
        vault=CredentialVault('test-slot',backend)
        self.assertIn('fictional-access-token',vault.read())
        saved=token_info();vault.save(saved);backend.set_password.assert_called_once_with(SERVICE,'test-slot',saved)
        vault.clear();backend.delete_password.assert_called_once_with(SERVICE,'test-slot')
        account=GoogleAccount('test',vault=MemoryVault(token_info()));account.disconnect()
        with self.assertRaisesRegex(GoogleError,'disconnected'):account.credentials()

    def test_no_plaintext_fallback_and_safe_errors(self):
        backend=Mock();backend.get_password.side_effect=RuntimeError('secret-person-and-token')
        with self.assertRaises(GoogleError) as error:CredentialVault('test',backend).read()
        self.assertNotIn('secret-person-and-token',str(error.exception))
        backend.get_password.return_value=None;backend.get_password.side_effect=None
        backend.set_password.side_effect=RuntimeError('secret')
        with self.assertRaisesRegex(GoogleError,'securely'):CredentialVault('test',backend).save('secret')
        with patch('keyring.get_keyring',return_value=Mock()),patch('google_account.os.name','posix'),patch('sys.platform','linux'):
            with self.assertRaisesRegex(GoogleError,'secure credential store'):CredentialVault('test')

    def test_refresh_stored_credentials_and_revocation(self):
        vault=MemoryVault(token_info(expired=True));account=GoogleAccount('test',vault=vault)
        def refreshed(cred, request):
            cred.token='fictional-new-token';cred.expiry=datetime.now(timezone.utc).replace(tzinfo=None)+timedelta(hours=1)
        with patch('google.oauth2.credentials.Credentials.refresh',autospec=True,side_effect=refreshed):
            self.assertEqual(account.credentials().token,'fictional-new-token')
        self.assertIn('fictional-new-token',vault.value)
        from google.auth.exceptions import RefreshError
        vault.value=token_info(expired=True)
        with patch('google.oauth2.credentials.Credentials.refresh',side_effect=RefreshError('private-token')):
            with self.assertRaisesRegex(GoogleError,'sign in again') as error:account.credentials()
            self.assertNotIn('private-token',str(error.exception))

    def test_missing_scopes_requires_reauthorization(self):
        info=json.loads(token_info());info['scopes']=['openid']
        account=GoogleAccount('test',vault=MemoryVault(json.dumps(info)))
        with self.assertRaisesRegex(GoogleError,'permissions have changed'):account.credentials()

    def test_desktop_client_only_and_official_endpoints(self):
        with tempfile.TemporaryDirectory() as tmp:
            path=Path(tmp)/'client.json'
            for bad in ({'type':'service_account'}, {'installed':dict(client_id='id',client_secret='secret',auth_uri='https://example.test',token_uri='https://oauth2.googleapis.com/token')}):
                path.write_text(json.dumps(bad))
                with self.assertRaisesRegex(GoogleError,'Desktop App'):client_configuration(path)

    def test_login_loopback_pkce_and_cancel_does_not_replace_token(self):
        vault=MemoryVault(token_info());account=GoogleAccount('test',vault=vault)
        cred=Mock();cred.refresh_token='refresh';cred.has_scopes.return_value=True;cred.granted_scopes=list(SCOPES);cred.to_json.return_value=token_info()
        flow=Mock();flow.run_local_server.return_value=cred
        with patch('google_account.client_configuration',return_value={}),patch('google_auth_oauthlib.flow.InstalledAppFlow.from_client_config',return_value=flow) as create,patch.object(account,'identity',return_value='staff@example.test'):
            self.assertEqual(account.connect(),'staff@example.test')
            self.assertTrue(create.call_args.kwargs['autogenerate_code_verifier'])
            self.assertEqual(flow.run_local_server.call_args.kwargs['host'],'127.0.0.1')
            self.assertEqual(flow.run_local_server.call_args.kwargs['port'],0)
            flow.run_local_server.side_effect=RuntimeError('cancelled-secret')
            saved=vault.value
            with self.assertRaises(GoogleError) as error:account.connect()
            self.assertEqual(vault.value,saved);self.assertNotIn('cancelled-secret',str(error.exception))

    def test_identity_failure_does_not_switch_saved_account(self):
        vault=MemoryVault(token_info());saved=vault.value;account=GoogleAccount('test',vault=vault)
        cred=Mock();cred.refresh_token='refresh';cred.has_scopes.return_value=True;cred.granted_scopes=list(SCOPES)
        flow=Mock();flow.run_local_server.return_value=cred
        with patch('google_account.client_configuration',return_value={}),patch('google_auth_oauthlib.flow.InstalledAppFlow.from_client_config',return_value=flow),patch.object(account,'identity',side_effect=GoogleError('Cannot check new account')):
            with self.assertRaisesRegex(GoogleError,'Cannot check new account'):account.connect()
        self.assertEqual(vault.value,saved)

    def test_partial_consent_is_not_saved(self):
        vault=MemoryVault();account=GoogleAccount('test',vault=vault)
        cred=Mock();cred.refresh_token='refresh';cred.has_scopes.return_value=True;cred.granted_scopes=['openid']
        flow=Mock();flow.run_local_server.return_value=cred
        with patch('google_account.client_configuration',return_value={}),patch('google_auth_oauthlib.flow.InstalledAppFlow.from_client_config',return_value=flow):
            with self.assertRaisesRegex(GoogleError,'declined'):account.connect()
        self.assertIsNone(vault.value)

    def test_drive_pagination_lists_metadata_only(self):
        account=GoogleAccount('test',vault=MemoryVault());service=Mock()
        service.files.return_value.list.return_value.execute.side_effect=[{'files':[{'id':'b','name':'B'}],'nextPageToken':'next'},{'files':[{'id':'a','name':'A'}]}]
        with patch.object(account,'service',return_value=service) as build:
            self.assertEqual([s['id'] for s in account.spreadsheets()],['a','b'])
            build.assert_called_once_with('drive','v3')
        calls=service.files.return_value.list.call_args_list
        self.assertEqual(calls[1].kwargs['pageToken'],'next')
        self.assertEqual(calls[0].kwargs['fields'],'nextPageToken,files(id,name)')

    def test_header_read_only_selected_sheet_and_escaped_title(self):
        account=GoogleAccount('test',vault=MemoryVault());service=Mock()
        service.spreadsheets.return_value.values.return_value.get.return_value.execute.return_value={'values':[['Name','Phone','ID']]}
        worksheet={'sheetId':7,'title':"Team's forms",'gridProperties':{'columnCount':1200}}
        with patch.object(account,'service',return_value=service):
            self.assertEqual(account.headers('selected-only',worksheet),['Name','Phone','ID'])
        call=service.spreadsheets.return_value.values.return_value.get.call_args.kwargs
        self.assertEqual(call['spreadsheetId'],'selected-only');self.assertTrue(call['range'].startswith("'Team''s forms'!A1:"));self.assertTrue(call['range'].endswith('1'))

    def test_gmail_transport_does_not_replay_broken_connection(self):
        from google_account import SingleAttemptHttp
        from http.client import RemoteDisconnected
        from requests.exceptions import ConnectionError
        cred=Mock()
        with patch('requests.Session') as session:
            transport=SingleAttemptHttp(cred)
            session.return_value.request.side_effect=ConnectionError(RemoteDisconnected('connection closed after submission'))
            with self.assertRaises(ConnectionError):
                transport.request('https://gmail.googleapis.com/gmail/v1/users/me/messages/send',method='POST',body='fictional')
            session.return_value.request.assert_called_once()
            self.assertFalse(session.return_value.request.call_args.kwargs['allow_redirects'])
            self.assertEqual(session.return_value.mount.call_args.args[1].max_retries.total,0)
            self.assertEqual(session.return_value.request.call_args.kwargs['timeout'],30)

    def test_real_gmail_client_uses_single_attempt_transport(self):
        from google_account import SingleAttemptHttp
        account=GoogleAccount('test',vault=MemoryVault(token_info()))
        service=account.service('gmail','v1')
        self.assertIsInstance(service._http,SingleAttemptHttp)
        response=Mock(status_code=200,headers={'content-type':'application/json'},content=b'{"id":"mock-id"}')
        with patch.object(service._http.session,'request',return_value=response) as request:
            result=service.users().messages().send(userId='me',body={'raw':'fictional'}).execute(num_retries=0)
            self.assertEqual(result['id'],'mock-id');request.assert_called_once()
            self.assertEqual(request.call_args.args[0],'POST')
        service.close()


class SourceTests(unittest.TestCase):
    def source(self):return dict(type='volunteer',spreadsheet='selected',worksheet_id=7,range="'Old title'!A1:Z",destination='Volunteers',name='Name',phone='Phone',date='Participation date',key='Submission ID')

    def test_required_mapping_and_timestamp_rejection(self):
        source=self.source();validate_source(source)
        for field in ('name','phone','date'):
            invalid=dict(source);invalid[field]=''
            with self.assertRaises(ValueError):validate_source(invalid)
        with self.assertRaisesRegex(ValueError,'timestamp'):validate_source(dict(source,date='Timestamp'))
        with self.assertRaisesRegex(ValueError,'unique registration ID'):validate_source(dict(source,key=''))
        validate_source(dict(source,key='',row_identity_ack=True))
        with self.assertRaisesRegex(ValueError,'email'):validate_source(dict(source,type='beneficiary'))

    def test_stable_sheet_id_updates_after_reorder_and_rename(self):
        source=self.source();service=Mock();api=service.spreadsheets.return_value
        api.get.return_value.execute.return_value={'sheets':[{'properties':{'sheetId':7,'title':'Renamed','gridProperties':{'columnCount':30}}}]}
        grid=[['Name','Phone','Participation date','Submission ID','Other answer'],['Alice','0123456789','2026-12-01','0001','First'],['Bob','0123456788','2026-12-02','0002','Second']]
        api.values.return_value.get.return_value.execute.return_value={'values':grid}
        first=fetch({'sources':[source]},service)
        with tempfile.TemporaryDirectory() as tmp:
            store=ExcelStore(Path(tmp)/'records.xlsx');self.assertEqual(store.merge_google(first),(2,0))
            old_ids={r['__fb_key']:r['__fb_id'] for r in store.table('Volunteers')[1]}
            grid[1][-1]='Changed';grid[1],grid[2]=grid[2],grid[1]
            second=fetch({'sources':[source]},service);self.assertEqual(store.merge_google(second),(0,1))
            self.assertEqual(old_ids,{r['__fb_key']:r['__fb_id'] for r in store.table('Volunteers')[1]})
            self.assertEqual(store.merge_google(second),(0,0))
        self.assertIn('gid=7',first[0]['source']);self.assertIn('Google · Other answer',first[0]['headers'])
        self.assertTrue(api.values.return_value.get.call_args.kwargs['range'].startswith("'Renamed'!"))

    def test_destination_case_change_cannot_create_duplicate_worksheets(self):
        with tempfile.TemporaryDirectory() as tmp:
            store=ExcelStore(Path(tmp)/'records.xlsx')
            batch={'sheet':'Volunteers','source':'source','headers':['Participant Name'],
                   'rows':[{'Participant Name':'Example','__fb_key':'one','__fb_source':'source'}]}
            store.merge_google([batch]);original=list(store.sheets)
            with self.assertRaisesRegex(ValueError,'capitalization'):
                store.merge_google([dict(batch,sheet='volunteers')])
            self.assertEqual(store.sheets,original)
            self.assertEqual(len(store.table('Volunteers')[1]),1)

    def test_duplicate_id_stops_whole_fetch(self):
        source=self.source();service=Mock();api=service.spreadsheets.return_value
        api.get.return_value.execute.return_value={'sheets':[{'properties':{'sheetId':7,'title':'Responses'}}]}
        api.values.return_value.get.return_value.execute.return_value={'values':[['Name','Phone','Participation date','Submission ID'],['Alice','0123456789','Monday','1'],['Bob','0123456788','Tuesday','1']]}
        with self.assertRaisesRegex(ValueError,'duplicated'):fetch({'sources':[source]},service)


if __name__=='__main__':unittest.main()
