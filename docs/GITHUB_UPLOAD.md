# GitHub에 소스 반영하기

대상 저장소: https://github.com/h19h29-design/rename

2026-09-13 전달 시점에는 `.gitignore` 초기 커밋 `c741ad90808a816bcd4ee09b93f3444339abe588`만 main에 반영되어 있습니다. 전체 소스 업로드는 연결 도구의 보안 확인 단계에서 차단되어 미완료입니다. ZIP에는 전체 소스가 들어 있습니다.

## 로컬 Git으로 반영

Git이 설치되고 본인의 GitHub 계정으로 인증된 터미널을 사용합니다. 다른 작업 폴더와 겹치지 않는 새 위치에서 다음 명령을 실행합니다.

```powershell
git clone https://github.com/h19h29-design/rename.git rename-github
cd rename-github
```

이 ZIP을 푼 `RENAME_source` 폴더의 **내용물**을 `rename-github` 안으로 복사합니다. `.git` 폴더는 지우거나 덮어쓰지 않습니다. 실제 처리 대상 문서나 결과 파일은 복사하지 마세요.

```powershell
git status --short
git add .gitignore rename_app tests docs .github run.py build_windows.py requirements.txt requirements-dev.txt START_WINDOWS.cmd BUILD_WINDOWS.cmd README.md THIRD_PARTY_NOTICES.md "먼저_읽어주세요.txt"
git diff --cached --stat
git commit -m "feat: add local Windows document pseudonymization utility"
git push origin main
```

`git push`가 충돌이나 권한 문제로 거부되면 강제 푸시하지 말고 원인을 확인합니다. 저장소가 그동안 변경되었다면 현재 변경사항과 합쳐야 합니다. 인증 토큰을 소스나 채팅에 붙여 넣지 마세요.

## Windows 배포본

포함된 `.github/workflows/windows.yml`은 main 반영 후 Windows 테스트 → EXE 패키징 → 패키징된 GUI 스모크 테스트 → Artifact/미리보기 Release 순서로 실행하도록 작성했습니다. 아직 실제 실행 결과는 확인하지 못했습니다. 저장소 설정에서 Actions 실행/워크플로 권한을 허용해야 할 수 있습니다.

실패하면 Actions 로그를 확인하고 수정합니다. 성공했다는 로그와 생성된 ZIP을 확인하기 전에는 EXE 배포 완료로 기록하지 마세요. 설치된 Excel·한글 연동은 CI와 별개로 실제 PC에서 검증해야 합니다.
