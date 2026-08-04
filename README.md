# AI Report Summary Project

AER Excel 테스트 레포트를 분석하여 Summary 시트를 자동 생성하는 Python 프로젝트입니다.

> 저장소는 회사 업무 자료와 샘플 레포트를 포함할 수 있으므로 **Private Repository**로 유지하세요.

## 현재 상태

- GitHub 표준 폴더 구조 구성 완료
- 프로젝트 명세와 인수인계 자료 포함
- 샘플 AER Excel 포함
- 최신 Python 소스는 아직 미포함
- 실제 사용 중인 `summary_generator.py` 및 관련 소스를 `src/`에 추가해야 함

## 주요 기능

- 대상 시트 자동 인식
- Major / Middle / Minor 계층 파싱
- Waveform O/X 및 파일 현황 집계
- 테스트 진행률 계산
- Daily Progress 생성
- Missing Image 검출
- Summary 시트 생성 및 스타일 적용
- 중복되지 않는 출력 파일명 생성

## 저장소 구조

```text
.github/        GitHub Issue, PR, Actions 설정
config/         설정 파일
docs/           프로젝트 문서와 명세
knowledge_base/ 확정 규칙과 설계 결정
prompts/        ChatGPT 작업 프롬프트
release/        배포 파일 및 릴리스 노트
samples/        입력/출력 샘플
src/            실제 Python 소스
tests/          테스트 및 회귀 검증
tools/          개발/점검 보조 도구
```

## 가장 먼저 할 일

1. 실제 사용 중인 최신 Python 파일을 `src/`에 업로드합니다.
2. `src/SOURCE_REQUIRED.md`를 확인합니다.
3. 샘플 Excel로 실행 결과를 검증합니다.
4. 검증 완료 후 `v0.1.0-baseline` 태그를 생성합니다.

## 새 ChatGPT 계정에서 시작하는 방법

1. 이 저장소의 `README.md`를 읽게 합니다.
2. `knowledge_base/PROJECT_DNA.md`와 `prompts/NEW_CHATGPT_ACCOUNT_PROMPT.md`를 제공합니다.
3. 수정할 최신 소스 파일을 함께 제공합니다.
4. 기존 기능을 유지하는 최소 수정 방식으로 작업합니다.
