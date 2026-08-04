# AI Report Summary Project

AER(실장 검토) Excel 레포트를 분석하여 테스트 진행 현황과 Summary 시트를 자동 생성하기 위한 Python 프로젝트입니다.

## 현재 저장소 상태

이 패키지는 **GitHub 이관용 기준 저장소**입니다. 프로젝트 규칙, 샘플 Excel, 인수인계 문서와 협업 템플릿은 포함되어 있으나, 현재 대화에 최신 Python 소스가 제공되지 않아 `src/`에는 소스 투입 위치만 준비되어 있습니다.

> 최신 `summary_generator.py` 및 관련 Python 파일을 `src/`에 추가한 뒤 첫 번째 기준 태그를 생성하십시오.

## 핵심 목표

- Waveform 및 테이블 데이터를 이용한 테스트 진행률 계산
- Summary 시트를 원본 Workbook의 마지막에 생성
- 전체/시트별 Daily Progress 집계
- 파일명은 있으나 Embedded Image가 없는 항목 검출
- 기존 병합 셀, 스타일과 레포트 구조 보존
- Drag & Drop, 진행률 표시, 단계별 로그 제공

## 빠른 시작

1. `docs/PROJECT_CONTEXT.md`를 읽습니다.
2. `docs/specs/REQUIREMENTS.md`와 `docs/specs/ALGORITHM_SPEC.md`를 확인합니다.
3. 최신 Python 소스를 `src/`에 넣습니다.
4. 원본 샘플은 `samples/`에서 확인합니다.
5. 코드 변경 전 `tests/REGRESSION_CHECKLIST.md`를 기준으로 영향 범위를 검토합니다.

## 대상 시트명 규칙

`<Operation Mode>_<Input><Ratio>`

- Operation Mode: `RVS`, `FWD`
- Input: `V` = VBUS, `W` = WPC
- Ratio: `1`~`4`
- 예: `RVS_V4`, `FWD_W2`

## 브랜치 운영

- `main`: 검증 완료된 배포 기준
- `develop`: 통합 개발
- `feature/<topic>`: 기능 단위 개발
- `fix/<issue>`: 버그 수정
- `docs/<topic>`: 문서 변경

자세한 방식은 `CONTRIBUTING.md`를 참고하십시오.
