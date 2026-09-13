"""Optional Windows-only legacy conversion, isolated from the GUI process."""
from __future__ import annotations
from pathlib import Path
import multiprocessing as mp
import sys
import tempfile
from .documents import DocumentError, MAX_FILE


def _convert_child(source: str, target: str, kind: str, pipe):
    app = book = None
    pythoncom = None
    try:
        import pythoncom
        from win32com.client import DispatchEx
        pythoncom.CoInitialize()
        if kind == '.xls':
            app = DispatchEx('Excel.Application')
            app.Visible = False
            app.DisplayAlerts = False
            app.EnableEvents = False
            app.AskToUpdateLinks = False
            app.AutomationSecurity = 3  # msoAutomationSecurityForceDisable BEFORE opening.
            book = app.Workbooks.Open(Filename=source,UpdateLinks=0,ReadOnly=True,
                Password='',WriteResPassword='',IgnoreReadOnlyRecommended=True,AddToMru=False)
            if int(book.HasVBProject):
                raise DocumentError('매크로가 있는 XLS는 처리하지 않습니다. 매크로 없는 XLSX로 저장해 주세요.')
            book.CheckCompatibility = False
            book.SaveAs(Filename=target,FileFormat=51,CreateBackup=False,AddToMru=False)
        else:
            app = DispatchEx('HWPFrame.HwpObject')
            # Do not install security-bypass DLLs or change the user's registry.
            # An installed Hancom edition may ask the user to approve local file access.
            try: app.XHwpWindows.Item(0).Visible = True
            except Exception: pass
            if not app.Open(source,'HWP',''):
                raise DocumentError('한글에서 HWP를 열지 못했습니다. 보호·암호 여부를 확인해 주세요.')
            if not app.SaveAs(target,'HWPX',''):
                raise DocumentError('설치된 한글에서 HWPX 저장에 실패했습니다. 직접 HWPX로 저장한 후 다시 넣어 주세요.')
        pipe.send('ok')
    except ImportError:
        pipe.send('윈도우 연동 모듈이 없습니다. START_WINDOWS.cmd로 실행하거나 배포본을 사용해 주세요.')
    except DocumentError as exc:
        pipe.send(str(exc))
    except Exception:
        pipe.send('설치된 Office 프로그램과 연동하지 못했습니다. Excel/한글 설치 및 로컬 파일 접근 승인 창을 확인해 주세요.')
    finally:
        try:
            if book is not None: book.Close(SaveChanges=False)
        except Exception: pass
        try:
            if app is not None: app.Quit()
        except Exception: pass
        if pythoncom is not None:
            try: pythoncom.CoUninitialize()
            except Exception: pass
        pipe.close()


def convert_legacy(source: Path) -> tuple[bytes,str]:
    if sys.platform != 'win32':
        raise DocumentError('XLS/HWP 변환은 Windows와 설치된 Excel/한글이 필요합니다. XLSX/HWPX로 저장한 파일은 별도 설치 없이 처리합니다.')
    target_suffix = '.xlsx' if source.suffix.lower()=='.xls' else '.hwpx'
    context = mp.get_context('spawn')
    with tempfile.TemporaryDirectory(prefix='rename-private-') as folder:
        target = Path(folder)/('conversion'+target_suffix)
        receiver,sender = context.Pipe(duplex=False)
        process = context.Process(target=_convert_child,args=(str(source),str(target),source.suffix.lower(),sender),daemon=True)
        process.start()
        sender.close()
        try:
            if not receiver.poll(120):
                raise DocumentError('Office 변환 응답이 없습니다. 한글의 접근 승인 창을 확인하거나 직접 XLSX/HWPX로 저장해 주세요.')
            try: response = receiver.recv()
            except EOFError:
                raise DocumentError('Office 변환 작업이 예기치 않게 종료되었습니다.') from None
            process.join(10)
            if response != 'ok': raise DocumentError(response)
            if not target.exists() or not 0 < target.stat().st_size <= MAX_FILE:
                raise DocumentError('변환 파일이 없거나 안전 크기를 초과했습니다.')
            return target.read_bytes(),target_suffix
        finally:
            receiver.close()
            if process.is_alive():
                process.terminate()
                process.join(5)
