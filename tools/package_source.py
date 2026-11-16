"""Build an allowlisted source ZIP, excluding live records and credentials."""
from pathlib import Path
import hashlib
import sys
import zipfile
ROOT=Path(__file__).resolve().parents[1]
FILES=['google_account.py','google_dialog.py','email_delivery.py','OPERATOR_GUIDE_ZH.md','applications.py','README.md','VALIDATION.md','main.py','storage.py','certificates.py','google_sync.py',
       'requirements.txt','run.bat','run.sh','.gitignore',
       'templates/volunteer_template.pptx','templates/PLACE_TEMPLATES_HERE.txt',
       'examples/volunteer_sample.pptx','examples/sample_records.xlsx','examples/interface.png',
       'examples/google_selection.png','examples/family_email_review.png','examples/family_confirmation_demo.pdf']


def build(output):
    files=[ROOT/f for f in FILES]
    for directory in ('tests','tools'):
        files.extend(sorted((ROOT/directory).glob('*.py')))
    output=Path(output);output.parent.mkdir(parents=True,exist_ok=True)
    manifest=[]
    with zipfile.ZipFile(output,'w',zipfile.ZIP_DEFLATED) as archive:
        for file in files:
            relative=file.relative_to(ROOT).as_posix()
            archive.write(file,'Foodbank-System/'+relative)
            manifest.append(hashlib.sha256(file.read_bytes()).hexdigest()+'  '+relative)
        archive.writestr('Foodbank-System/SHA256SUMS.txt','\n'.join(manifest)+'\n')
    with zipfile.ZipFile(output) as archive:
        assert archive.testzip() is None
        assert not any('/data/' in n or '/.git/' in n or 'service-account' in n or 'record.xlsx' in n or 'local_settings.json' in n or 'oauth_client' in n or 'client_secret' in n or 'token.json' in n or n.endswith('.sqlite3') for n in archive.namelist())
    print(f'{output}: {len(files)} source/assets files, {output.stat().st_size} bytes')


if __name__=='__main__':build(sys.argv[1] if len(sys.argv)>1 else ROOT.parent/'Foodbank-System-source.zip')
