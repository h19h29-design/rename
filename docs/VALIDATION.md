# 검증 기록 — 2026-09-13

## 실행 환경과 결과

### Windows 11 x64 (2026-09-13)

Windows 11 x64 build 10.0.26200 / Python 3.12.10 / lxml 6.1.3 / pywin32 312 / PyInstaller 6.22.3에서 확인했습니다.

- 자동 테스트: **39 passed, 1 skipped**.
- 소스 GUI 스모크와 PyInstaller 패키징 실행 파일 GUI 스모크 통과.
- 로컬 포터블 ZIP 생성: `dist/RENAME-Windows-x64.zip`. SHA256: `cc2f9a95c550c0d5eb4bd283118c12f1f1922da12b808a4631958ba86ac28fe5`.
- Excel 16.0: 합성 XLSX 생성 → RE:NAME 처리 → Excel에서 열기/저장/재열기 시 시트 1개, 병합 A5:B5, 숨김 행 3, A2/B2 가명화가 유지됨을 확인.
- Excel로 만든 합성 XLS: pywin32 변환 경로로 XLSX 변환·열기 성공, 값과 병합 셀 유지 확인.
- Hancom COM `HWPFrame.HwpObject`가 등록되어 있지 않아 실제 HWP 변환과 HWPX 레이아웃은 미검증.

### GitHub 푸시·CI·릴리스 (2026-09-13)

- main 소스 커밋 `e758d9acfa053e96d6333bd981371735b8af862e`를 강제 푸시 없이 반영.
- GitHub Actions 실행 34745009276(run #1) 성공. 테스트·빌드·패키징된 GUI 스모크·Artifact 업로드·미리보기 릴리스 단계 모두 성공.
- 미리보기 릴리스 `v0.1.0-build.1.1`: https://github.com/h19h29-design/rename/releases/tag/v0.1.0-build.1.1
- 다운로드 ZIP: https://github.com/h19h29-design/rename/releases/download/v0.1.0-build.1.1/RENAME-Windows-x64.zip
- 릴리스 ZIP SHA256 `d8a5150d9f12fc74b5f6ef04b779ffbb8b7f947f7887f15cce59b1272fcdbcff` — 다운로드한 ZIP, `SHA256SUMS.txt`, GitHub 자산 digest와 일치.
- 다운로드한 릴리스 EXE의 자체 테스트와 GUI 스모크는 종료 코드 0.
- 로컬 빌드 ZIP SHA256 `cc2f9a95c550c0d5eb4bd283118c12f1f1922da12b808a4631958ba86ac28fe5`. 로컬과 GitHub 해시는 별도 빌드라 다르며, 다운로드에는 릴리스 해시를 사용.
- Artifact ID 10313228795 존재, 2026-10-13 만료. 사용자 안내에는 릴리스 URL을 우선 사용.

### 테스트 범위 (Linux / Python 3.13 / lxml 6.1.1 / Tk 8.6 / Xvfb 포함)

- 규칙 테스트: 이름·주민번호·생년월일·금액·일반 날짜, 날짜 유효성, 우선순위, 옵션 해제, 반복 치환.
- 합성 XLSX: 숨김 행, 수식 캐시, 공유 문자열 미사용 항목 제거, 병합 유지, 원본 SHA256 불변.
- 합성 HWPX: 여러 글자 서식에 나뉜 문자열 치환, 반복 이름, 미리보기/작성자 속성 제거.
- 추가 경계: 띄어 쓴 이름, 숫자의 지수 표기, 선행 0 복원, 대체값의 수식 주입 방지, 선택 원문 잔존 시 저장 중단.
- fail-closed: 외부 연결·쿼리·DDE·컨트롤 속성·매크로·대화상자 시트 XLSX 파트 차단, 경로 순회·DTD/엔터티·불투명한 첨부·알 수 없는 파일 형식 차단.
- 결과 저장: 반복 저장 시 새 폴더 사용, 미리보기 후 원본 변경 감지(최대 파일 크기만큼만 다시 읽는 bounded 재검사), 원문 없는 개수 보고서, 이미지 확인 없이는 저장 중단.
- UI: 실제 Tk 창 생성, 기본 옵션, 탐지 결과 표 연결, 설정 변경 시 재탐지 강제, 체크 해제, 원문 가리기.

기존 Linux 자동 테스트는 **33 passed**였고, 2026-09-13 Windows 재실행 기준은 **39 passed, 1 skipped**입니다.

## 아직 실제 사용자 PC에서 확인할 항목

1. 설치된 한글을 통한 HWP → HWPX 변환 및 파일 접근 승인 창 동작(`HWPFrame.HwpObject` 미등록).
2. 실제 한글에서 HWPX 결과 열림, 표·머리말·꼬리말·페이지 레이아웃 확인.
3. 실제 업무 서식에서 이름/생년월일 열 탐지와 누락·오탐 검토.
4. 이미지·도장·서명·스캔본과 자동 탐지 범위 밖 정보 직접 확인.
5. Windows 10, 다른 Office/한글 버전, 기관 정책 환경에서 재현 확인.

main 소스 푸시·Actions·미리보기 릴리스는 완료했습니다. 다만 Windows CI가 성공해도 Excel / Hancom 설치 연동 테스트가 자동으로 완료되는 것은 아닙니다. GitHub Windows 러너에는 기관의 실제 Office 환경이 없습니다.
