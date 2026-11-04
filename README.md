# 시험 성적 분석 보고서

XLS·XLSX·CSV 시험 성적을 검증하고 그래프가 포함된 간결한 한국어 Word 보고서를 만드는 Codex 스킬입니다.

## 기능

- 전체·반별 점수 분포와 문항별 변동성·난이도·변별력 분석
- 반 정보가 있으면 반×문항 히트맵 포함
- 공식 배점이 있으면 평균 득점률, 없으면 점수 획득률 사용
- 실제 0점과 결측 구분, 원본 총점 대조 및 독립 재계산
- Word 보고서와 300dpi 그래프 ZIP 생성

## 설치와 사용

이 저장소를 개인 Codex 스킬 폴더의 `exam-results-report`에 복제합니다. 기본 위치를 사용하는 경우:

```sh
git clone https://github.com/hyennnnnnn/exam-results-report.git ~/.codex/skills/exam-results-report
```

같은 이름의 스킬이 이미 설치되어 있으면 기존 폴더를 덮어쓰지 말고 업데이트 방법을 확인하세요. `CODEX_HOME`을 별도로 설정했다면 그 아래의 `skills` 폴더를 사용합니다. 비공개 저장소는 접근 권한이 있는 GitHub 계정으로 인증해야 합니다.

성적 파일을 첨부한 뒤 다음과 같이 요청합니다.

> $exam-results-report로 시험 결과를 정리해줘.

문항 배점표와 참고 Word 문서가 있으면 함께 제공할 수 있습니다. 실제 학생 데이터는 이 저장소에 포함되어 있지 않습니다.

## 실행 환경

Python 3.10 이상과 한글 글꼴이 필요합니다. 라이브러리 목록은 `scripts/requirements.txt`에 있습니다. 별도 환경에서 실행할 때는 작업용 가상환경에 설치하세요.

```sh
python -m venv .venv
source .venv/bin/activate
python -m pip install -r scripts/requirements.txt
python scripts/test_analysis.py
python scripts/test_report.py
```

두 번째 테스트는 가상 데이터로 문서와 그래프를 생성하며 완료 후 임시 파일을 정리합니다. 페이지 검토용 파일을 남기려면 `--out work/report-check`를 지정합니다. Word의 페이지 렌더링과 시각 검토는 별도의 문서 도구로 수행합니다.

## 분석 범위

기본 분석은 문항 점수가 모두 확인된 학생을 대상으로 합니다. 전 문항 결측과 일부 결측은 구분하여 집계하고, 실제 0점은 포함합니다. 원본과 문항 합계가 다르면 원인을 확인하기 전 진행하지 않습니다. 관측 최고점을 공식 배점으로 추정하지 않습니다.

총점만 있는 자료, 가중 합산, 음수 감점은 기본 분석 스크립트의 자동 처리 범위에 포함되지 않습니다. 스킬이 자료 구조와 시험 규칙을 확인하여 분석 범위를 조정합니다. 반별 차이의 유의성 검정은 자동으로 수행하지 않습니다.

## 구성

- `SKILL.md`: 스킬 진입점과 작업 흐름
- `references/`: 입력 설정, 분석 기준, 보고서 서식
- `scripts/analyze_exam.py`: 입력 조사와 점수 분석
- `scripts/build_report.py`: 그래프와 Word 생성
- `scripts/test_analysis.py`: 13개 가상 데이터 계산 테스트
- `scripts/test_report.py`: 히트맵 포함·제외 및 동일 점수 시험의 생성 검증

점수 파일과 보고서는 작업 폴더에서 관리하세요. `.gitignore`는 일반적인 성적 파일과 생성 결과물의 실수 업로드를 줄이기 위한 기본 설정입니다.
