# Codex 인수인계 — RE:NAME

다음 내용을 ZIP을 푼 소스 폴더를 연 Codex에 전달합니다.

---

현재 폴더에 구현된 RE:NAME을 이어서 검토하고 Windows 배포본을 완성해라. 새로 만들거나 기존 기능을 제거하지 마라.

사용자 요청: Windows에서 엑셀/한글을 선택하여 이름·주민등록번호·생년월일을 삭제 또는 대체. 금액·일반 날짜는 선택 옵션. 대상 저장소는 `https://github.com/h19h29-design/rename`이다.

현재 상태:
- Python/Tkinter UI, lxml ZIP/XML 처리, Windows pywin32 구형 파일 변환 어댑터가 구현되어 있다.
- `python -m pytest -q`는 Linux/Xvfb에서 33 passed였다.
- 전체 소스는 이 폴더에 있다. GitHub 연결 도구에서 소스 업로드가 차단되어 전체 반영하지 못했다. 저장소는 초기 .gitignore 커밋 c741ad9 상태였지만 현재 상태를 다시 읽어 확인하라.
- EXE는 아직 만들지 못했다. BUILD_WINDOWS.cmd, build_windows.py, .github/workflows/windows.yml은 작성되어 있지만 Windows 빌드는 미검증이다.
- Windows/Excel/Hancom 실제 파일 열림과 연동 검증은 미완료다.

실행 순서:
1. README.md, docs/VALIDATION.md, docs/IMPLEMENTATION_PLAN.md, 전체 코드를 읽고 git 상태를 확인한다. 기존 사용자 변경은 보존한다.
2. Windows라면 Python 3.12 x64 가상환경에서 requirements-dev.txt를 설치하고 테스트를 실행한다. Mac/Linux에서는 변환 핵심 테스트와 Windows 빌드를 구분한다.
3. 공개 저장소에 실제 원본 문서·대응표·출력·인증정보가 들어가지 않는지 확인한다. 테스트는 합성 자료만 사용한다.
4. 원격 저장소를 정상 clone/fetch하고 이 소스를 정식 커밋·푸시한다. 원격 변경이 있으면 병합 검토하며 강제 푸시는 하지 않는다.
5. Windows에서 build_windows.py를 실행하거나 Actions 결과를 확인한다. 실패 시 실제 로그를 바탕으로 수정한다. EXE와 _internal을 포함하는 ZIP이 필요하다.
6. 실제 Windows에서 익명 데이터 XLSX/HWPX를 열어 저장·재열기·레이아웃을 검증한다. 설치된 Excel이 있으면 XLS, 한글이 있으면 HWP→HWPX 및 파일 접근 승인 창을 검증한다. Office가 없으면 해당 검증을 미완료로 명시한다.
7. 검증 기록·최종 SHA·EXE 다운로드 위치·미검증 항목을 보고한다.

지켜야 할 동작:
- 원본 덮어쓰기 금지, 외부 API/서버로 문서 전송 금지, 원문 대응표/원본 파일명 로그 저장 금지.
- 자동 탐지 후 사용자가 원문·대체값·위치·선택 여부를 검토한다. 금액·일반 날짜는 기본 OFF.
- 자동 이름 인식의 한계를 숨기지 말고 직접 이름 추가를 유지한다.
- 이미지·도장·서명은 OCR/삭제하지 않으며 별도 확인을 강제한다.
- XLSX 결과는 공유용 정적 복사본이다. 수식 캐시가 없는 셀은 비워진다는 경고를 유지한다.
- 지원 불가 차트·피벗·매크로·첨부·변경 이력은 차단하며 안전하다고 오인시키지 않는다.
- HWP는 설치된 한글을 통해 HWPX로 변환한다. 레지스트리/전역 보안 우회는 하지 않는다.
- 완전한 익명성/모든 개인정보 제거/Windows 테스트 성공을 근거 없이 보장하지 않는다.

---
