# GitHub 반영 상태

대상 저장소: https://github.com/h19h29-design/rename

**2026-09-13 기준 소스 업로드와 미리보기 릴리스가 완료되었습니다.** 전체 소스는 `main`의 커밋 `e758d9acfa053e96d6333bd981371735b8af862e`로 강제 푸시 없이 반영되었습니다.

## 완료된 결과

- GitHub Actions 실행 34745009276(run #1) 성공: 테스트 → EXE 패키징 → 패키징된 GUI 스모크 → Artifact 업로드 → 미리보기 릴리스 단계 모두 성공.
- 미리보기 릴리스 `v0.1.0-build.1.1`: https://github.com/h19h29-design/rename/releases/tag/v0.1.0-build.1.1
- 다운로드: https://github.com/h19h29-design/rename/releases/download/v0.1.0-build.1.1/RENAME-Windows-x64.zip
- 릴리스 ZIP SHA256 `d8a5150d9f12fc74b5f6ef04b779ffbb8b7f947f7887f15cce59b1272fcdbcff`는 다운로드한 ZIP, `SHA256SUMS.txt`, GitHub 자산 digest와 일치합니다.
- Artifact ID 10313228795도 존재하지만 2026-10-13에 만료되므로 사용자 안내에는 릴리스 URL을 우선 사용합니다.

## 향후 변경을 반영할 때

1. 변경 전에 원격을 fetch하고 현재 상태를 확인합니다.
2. 초기 `.gitignore` 커밋 `c741ad90808a816bcd4ee09b93f3444339abe588`는 이미 반영되어 있으므로 다시 만들지 않습니다.
3. 실제 처리 대상 문서·결과 파일·대응표·인증정보는 커밋하지 마세요. 테스트는 합성 자료만 사용합니다.
4. 기존 방식대로 커밋·푸시합니다. `git push`가 충돌이나 권한 문제로 거부되면 **강제 푸시하지 말고** 원인을 확인합니다. 원격이 그동안 변경되었다면 현재 변경사항과 합치는 방법을 검토합니다.
5. 인증 토큰을 소스나 채팅에 붙여 넣지 마세요.

## Windows 배포본

포함된 `.github/workflows/windows.yml`은 main 반영 후 Windows 테스트 → EXE 패키징 → 패키징된 GUI 스모크 테스트 → Artifact/미리보기 Release 순서로 실행되며, 첫 실행(run 34745009276)은 성공했습니다. 저장소 설정에서 Actions 실행/워크플로 권한이 필요할 수 있습니다.

Windows CI가 성공해도 설치된 Excel·한글 연동 검증이 자동으로 완료되는 것은 아닙니다. 실제 PC에서 별도로 확인해야 합니다.
