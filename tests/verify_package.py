from pathlib import Path
import hashlib
import json
import subprocess
import zipfile

ROOT=Path(__file__).resolve().parent.parent
archives = sorted(ROOT.parent.glob('OfficePDF_*_Windows11_*.zip'))
if not archives:
    raise FileNotFoundError("Cannot find package zip in " + str(ROOT.parent))
archive = archives[-1]
product = ROOT / '成品'
destination = ROOT / 'tests' / 'temp_extract_check'
if destination.exists():
    import shutil
    shutil.rmtree(destination)

verify_icon_exe = ROOT / 'tests' / 'VerifyIcon.exe'
if not verify_icon_exe.exists():
    csc = Path(r"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe")
    if csc.exists():
        subprocess.run([str(csc), "/nologo", f"/out:{verify_icon_exe}", str(ROOT / 'tests' / 'VerifyIcon.cs')], check=True)

try:
    with zipfile.ZipFile(archive) as package:
        manifest = json.loads(package.read('SHA256.json').decode('utf-8-sig'))
        for item in manifest:
            data = package.read(item['File'].replace('\\', '/'))
            assert hashlib.sha256(data).hexdigest() == item['SHA256']
            assert len(data) == item['Bytes'] and data == (product / item['File']).read_bytes()
        assert set(package.namelist()) == {item['File'].replace('\\', '/') for item in manifest} | {'SHA256.json'}
        package.extractall(destination)

    exes = list(destination.glob('Office资料一键合并PDF_*.exe'))
    assert exes, "No GUI executable found in package"
    exe = exes[0]

    native_icons_dir = ROOT / 'tests' / 'temp_native_icons'
    if verify_icon_exe.exists():
        result = subprocess.run([str(verify_icon_exe), str(exe), str(native_icons_dir)], capture_output=True, timeout=25)
        log = result.stdout.decode('utf-8', errors='replace') + '\n' + result.stderr.decode('utf-8', errors='replace')
        assert result.returncode == 0, log

    report = {'passed': True, 'files_verified': len(manifest), 'archive_sha256': hashlib.sha256(archive.read_bytes()).hexdigest(), 'extracted_gui_and_shell_icons': True}
    print(json.dumps(report, ensure_ascii=False, indent=2))
finally:
    import shutil
    if destination.exists():
        shutil.rmtree(destination, ignore_errors=True)
    temp_icons = ROOT / 'tests' / 'temp_native_icons'
    if temp_icons.exists():
        shutil.rmtree(temp_icons, ignore_errors=True)
    if verify_icon_exe.exists():
        try:
            verify_icon_exe.unlink()
        except Exception:
            pass

