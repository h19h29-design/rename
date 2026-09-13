# Codex 인수인계 — RE:NAME (완료 상태)

RE:NAME은 Windows에서 엑셀/한글 문서를 선택해 이름·주민등록번호·생년월일을 검토 후 대체·삭제하는 프로그램입니다. 금액·일반 날짜는 선택 옵션입니다. 대상 저장소: `https://github.com/h19h29-design/rename`.

## 완료된 상태 (2026-09-13)

- Python/Tkinter UI, lxml ZIP/XML 처리, Windows pywin32 구형 파일 변환 어댑터 구현 완료.
- 전체 소스가 main 커밋 `e758d9acfa053e96d6333bd981371735b8af862e`로 강제 푸시 없이 반영됨.
- GitHub Actions 실행 34745009276(run #1) 성공: 테스트 → EXE 패키징 → 패키징된 GUI 스모크 → Artifact 업로드 → 미리보기 릴리스 단계 모두 성공.
- 미리보기 릴리스 `v0.1.0-build.1.1`: https://github.com/h19h29-design/rename/releases/tag/v0.1.0-build.1.1
- 다운로드: https://github.com/h19h29-design/rename/releases/download/v0.1.0-build.1.1/RENAME-Windows-x64.zip
- 릴리스 ZIP SHA256 `d8a5150d9f12fc74b5f6ef04b779ffbb8b7f947f7887f15cce59b1272fcdbcff`은 다운로드 ZIP·`SHA256SUMS.txt`·GitHub 자산 digest와 일치. 다운로드한 릴리스 EXE 자체 테스트와 GUI 스모크는 종료 코드 0.
- 로컬 빌드 ZIP SHA256은 `cc2f9a95c550c0d5eb4bd283118c12f1f1922da12b808a4631958ba86ac28fe5`이며 별도 빌드라 릴리스와 다름. 다운로드에는 릴리스 해시를 사용.
- Windows 11 x64에서 자동 테스트 39 passed, 1 skipped.

## 남은 수동 검증 항목

1. 설치된 한글을 통한 HWP → HWPX 변환 및 파일 접근 승인 창 동작(`HWPFrame.HwpObject` 미등록).
2. 실제 한글에서 HWPX 결과 열림, 표·머리말·꼬리말·페이지 레이아웃 확인.
3. 실제 업무 서식에서 이름/생년월일 열 탐지와 누락·오탐 검토.
4. 이미지·도장·서명·스캔본과 자동 탐지 범위 밖 정보 직접 확인.
5. Windows 10, 다른 Office/한글 버전, 기관 정책 환경에서 재현 확인.

## 유지해야 할 동작

- 원본 덮어쓰기 금지, 외부 API/서버로 문서 전송 금지, 원문 대응표/원본 파일명 로그 저장 금지.
- 자동 탐지 후 사용자가 원문·대체값·위치·선택 여부를 검토. 금액·일반 날짜는 기본 OFF.
- 자동 이름 인식의 한계를 숨기지 말고 직접 이름 추가를 유지.
- 이미지·도장·서명은 OCR/삭제하지 않으며 별도 확인을 강제.
- XLSX 결과는 공유용 정적 복사본이며 수식 캐시가 없는 셀은 비워진다는 경고를 유지.
- 지원 불가 차트·피벗·매크로·첨부·변경 이력은 차단하며 안전하다고 오인시키지 않음.
- HWP는 설치된 한글을 통해 HWPX로 변환. 레지스트리/전역 보안 우회 금지.
- 완전한 익명성/모든 개인정보 제거를 근거 없이 보장하지 않음.
- 향후 변경 시 강제 푸시 금지.
