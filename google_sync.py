"""Read-only Google Sheets connection; no credentials are bundled."""
import re
import os
from pathlib import Path
from storage import text


def spreadsheet_id(value):
    if '/forms/' in value:
        raise ValueError('This is a Google Form link. Open the form Responses tab, click the Google Sheets icon, and copy that response spreadsheet URL.')
    hit=re.search(r'/spreadsheets/d/([A-Za-z0-9_-]+)',value)
    result=hit.group(1) if hit else value.strip()
    if not re.fullmatch(r'[A-Za-z0-9_-]+',result): raise ValueError('Enter a Google Sheets URL or spreadsheet ID.')
    return result


def fetch(config, service=None):
    sources=config.get('sources',[])
    if not sources: return []
    if service is None:
        from google.oauth2.service_account import Credentials
        from googleapiclient.discovery import build
        from google_auth_httplib2 import AuthorizedHttp
        import httplib2
        path=Path(config.get('credentials',''))
        if not path.is_file():raise ValueError('Select your Google service-account JSON in Google Sheets Settings.')
        cred=Credentials.from_service_account_file(str(path),scopes=['https://www.googleapis.com/auth/spreadsheets.readonly'])
        service=build('sheets','v4',http=AuthorizedHttp(cred,http=httplib2.Http(timeout=30,ca_certs=os.environ.get('SSL_CERT_FILE') or os.environ.get('REQUESTS_CA_BUNDLE'))),cache_discovery=False)
    batches=[]; destinations=set()
    for source in sources:
        if source.get('type')=='donor':continue  # Retired workflow; preserve existing local history.
        sid=spreadsheet_id(source['spreadsheet'])
        title=source['destination'].strip()
        if not title or len(title)>31 or re.search(r'[\\/*?:\[\]]',title):raise ValueError('Destination worksheet names must be 1–31 characters without \\ / * ? : [ ].')
        if title in destinations: raise ValueError('Each Google source needs a different destination worksheet.')
        destinations.add(title)
        range_=source.get('range','').strip()
        if not range_:raise ValueError('Enter a range starting at the header row, e.g. Form Responses 1!A1:ZZ.')
        # Range must start with the headings; absolute row numbers serve as fallback keys.
        address=range_.rsplit('!',1)[-1]
        match=re.match(r'\$?[A-Za-z]+\$?(\d+)?(?:[:].*)?$',address)
        if not match:raise ValueError('Use an A1 cell range including its header row, e.g. Sheet1!A1:ZZ.')
        header_row=int(match.group(1) or 1)
        grid=service.spreadsheets().values().get(spreadsheetId=sid,range=range_,valueRenderOption='FORMATTED_VALUE').execute(num_retries=2).get('values',[])
        if not grid:continue
        original=[text(x).strip() for x in grid[0]]
        if not original or any(not h for h in original) or len(set(original))!=len(original):
            raise ValueError(f'{title}: Google header labels must be nonempty and unique.')
        if any(h.startswith('__fb_') for h in original):raise ValueError('Google headings cannot use reserved __fb_ names.')
        for field in ('name','phone','email','date','key'):
            col=source.get(field,'').strip()
            if col and col not in original: raise ValueError(f'{title}: column {col!r} was not found; use its exact header.')
        if not source.get('name'): raise ValueError(f'{title}: choose the participant name column.')
        # Original form columns retained; standard columns provide common mappings.
        standard=['Participant Name','Type','Phone','Email','Date']
        extras=[f'Google · {h}' for h in original]
        headers=standard+extras
        source_id=f'{sid}:{range_.split("!")[0] if "!" in range_ else "default"}:{source["type"]}:key={source.get("key") or "row-number"}'
        rows=[]; seen=set()
        for offset,line in enumerate(grid[1:],1):
            if not any(text(v).strip() for v in line):continue
            raw={h:text(line[i]) if i<len(line) else '' for i,h in enumerate(original)}
            key=raw.get(source.get('key',''),'') if source.get('key') else f'row:{header_row+offset}'
            if not key or key in seen:raise ValueError(f'{title}: response IDs are blank or duplicated. Use a unique response ID column.')
            seen.add(key)
            item={'Participant Name':raw.get(source.get('name'),''),'Type':source['type'],'Phone':raw.get(source.get('phone'),''),'Email':raw.get(source.get('email'),''),'Date':raw.get(source.get('date'),'')}
            item.update({f'Google · {h}':raw[h] for h in original})
            item.update(__fb_source=source_id,__fb_key=key)
            rows.append(item)
        batches.append({'sheet':title,'source':source_id,'headers':headers,'rows':rows})
    return batches
