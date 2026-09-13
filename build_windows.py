"""Build, test and package on Windows. Does not sign the executable."""
from pathlib import Path
import hashlib
import shutil
import subprocess
import sys
import importlib.metadata

ROOT = Path(__file__).resolve().parent

def run(*args):
    subprocess.run([sys.executable,*args],cwd=ROOT,check=True)

def main():
    if sys.platform != 'win32':
        raise SystemExit('Build this Windows executable on Windows, or use the GitHub Actions workflow.')
    if sys.maxsize <= 2**32:
        raise SystemExit('Use 64-bit Python.')
    run('-m','pytest','-q')
    run('run.py','--self-test','--gui-smoke')
    run('-m','PyInstaller','--noconfirm','--clean','--onedir','--windowed','--name','RENAME',
        '--hidden-import','win32timezone','--hidden-import','win32com.client','--hidden-import','pythoncom',
        '--collect-all','lxml','run.py')
    exe=ROOT/'dist'/'RENAME'/'RENAME.exe'
    subprocess.run([str(exe),'--self-test','--gui-smoke'],cwd=exe.parent,check=True,timeout=45)
    from rename_app.gui import HELP
    (exe.parent/'사용안내.txt').write_text(HELP,encoding='utf-8-sig')
    shutil.copy2(ROOT/'README.md',exe.parent/'README.md')
    shutil.copy2(ROOT/'THIRD_PARTY_NOTICES.md',exe.parent/'THIRD_PARTY_NOTICES.md')
    notices=exe.parent/'licenses'
    notices.mkdir(exist_ok=True)
    for package in ('lxml','pywin32','pyinstaller'):
        dist=importlib.metadata.distribution(package)
        for file in dist.files or []:
            if any(x in str(file).lower() for x in ('license','copying')):
                source=Path(dist.locate_file(file))
                if source.is_file():
                    safe_parts=[p for p in Path(str(file)).parts if p not in {'.','..'} and ':' not in p and p not in {'/','\\'}]
                    dest=notices/package/Path(*safe_parts)
                    dest.parent.mkdir(parents=True,exist_ok=True)
                    shutil.copy2(source,dest)
    python_license=Path(sys.base_prefix)/'LICENSE.txt'
    if python_license.exists(): shutil.copy2(python_license,notices/'PYTHON-LICENSE.txt')
    for license_file in (Path(sys.base_prefix)/'tcl').rglob('license.terms'):
        shutil.copy2(license_file,notices/f'{license_file.parent.name}-license.terms')
    (exe.parent/'BUILD_INFO.txt').write_text(
        f'Python: {sys.version}\nlxml: {importlib.metadata.version("lxml")}\n'
        f'pywin32: {importlib.metadata.version("pywin32")}\n'
        f'PyInstaller: {importlib.metadata.version("pyinstaller")}\n',encoding='utf-8')
    archive=Path(shutil.make_archive(str(ROOT/'dist'/'RENAME-Windows-x64'),'zip',ROOT/'dist','RENAME'))
    digest=hashlib.sha256(archive.read_bytes()).hexdigest()
    (ROOT/'dist'/'SHA256SUMS.txt').write_text(f'{digest}  {archive.name}\n',encoding='ascii')
    print(f'Built: {archive.name}\nSHA256: {digest}')

if __name__=='__main__': main()
