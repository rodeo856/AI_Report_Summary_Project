# GitHub 업로드 방법

현재 저장소에 README.md만 있는 상태를 기준으로 합니다.

1. 이 ZIP 파일을 PC에서 압축 해제합니다.
2. GitHub 저장소의 `Add file` → `Upload files`를 누릅니다.
3. 압축을 푼 폴더 **자체가 아니라 폴더 안의 모든 항목**을 선택해 드래그합니다.
4. 기존 README.md가 있다는 경고가 나오면 이 패키지의 README.md로 교체합니다.
5. Commit message에 `chore: initialize project structure`를 입력합니다.
6. `Commit changes`를 누릅니다.
7. 업로드 후 `src/`에 실제 최신 Python 소스를 추가합니다.

중요: ZIP 파일 자체를 Repository에 올리지 마세요.
