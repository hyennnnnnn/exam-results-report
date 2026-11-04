# 입력 설정과 실행

기본 파이프라인은 문항별 숫자 점수가 있는 XLS·XLSX·CSV에 사용한다. 총점만 있거나 가중 합산/음수 감점 시험이면 기준을 확인하고 별도 처리한다. 스크립트가 지원하지 않는 구조를 지원한다고 주장하지 않는다.

`inspect`는 최대 8행×12열의 자료를 출력한다. 개인 정보가 포함될 수 있으므로 작업 내부에서만 사용한다. 자동으로 학생 행이나 반을 확정하지 않는다. 병합 머리글, 다중 시트, 숨겨진 행, 공식 배점은 원본을 추가로 조사한다.

```sh
python SKILL_DIR/scripts/analyze_exam.py inspect scores.xlsx
python SKILL_DIR/scripts/analyze_exam.py analyze work/config.json --out work/analysis.json
python SKILL_DIR/scripts/build_report.py --analysis work/analysis.json --narrative work/narrative.json --out-dir outputs/시험보고서
```

설정 예시는 **가상 3문항 시험**이다. 행 번호는 1부터, 열은 Excel 문자다. 파일 상대 경로는 설정 파일의 부모 기준이다. `rows`는 학생 행을 정확히 지정한다. 같은 시험의 여러 반 시트는 sources에 각각 추가하며 items 순서를 동일하게 맞춘다. CSV는 기본 UTF-8-SIG이며 encoding과 delimiter를 설정할 수 있다.

```json
{
  "title": "예시 중간고사",
  "sources": [{
    "path": "../scores.xlsx", "sheet": "성적",
    "rows": [2, 3, 4, 5],
    "item_columns": ["D", "E", "F"],
    "id_column": "A", "class_column": "B", "total_column": "G"
  }],
  "items": [{"label": "1", "max": 2}, {"label": "2", "max": 3}, {"label": "3", "max": 5}],
  "total_max": 10,
  "max_source": "시험지의 공식 배점표",
  "missing_tokens": ["미입력"],
  "context_notes": ["추가로 확인된 시험별 조건만 기록합니다."]
}
```

`rows` 대신 `first_row`와 `last_row`(양끝 포함)를 지정할 수 있다. 둘을 같이 지정하지 않는다. 학생 아닌 행을 포함하지 않는다. `exclude_rows`는 `{ "7": "원본에 확인된 취소 기록" }`처럼 근거를 기록한다. 반 열 대신 `class_label`을 쓸 수 있으나 확인된 실제 반 이름이어야 한다. 둘 다 지정하지 않는다. class_label이 없으면 시트명을 자동으로 반으로 쓰지 않는다. sheet는 XLS/XLSX에 필수다. 중복 ID는 전체 sources에서 검사하므로 학생 ID가 반마다 재사용되는 출석번호라면 이를 id_column으로 쓰지 않는다. ID가 없으면 중복 학생 확인이 제한됨을 보고서에 적는다.

`max`와 `total_max`는 모르면 생략/null. 하나라도 알려진 배점이 있으면 `max_source`에 근거를 적는다. 전체 문항 배점이 알려지고 total_max가 없으면 그 합을 공식 총점으로 사용한다. 일부 배점만 있으면 모든 문항을 같은 척도로 비교하기 위해 획득률·원점수 차이를 쓴다. 확인된 배점은 해당 문항의 점수 범위 검증에 계속 사용한다.

`analysis.json`에는 원본 경로와 익명 점수 행렬이 들어 있어 작업용으로만 보관한다. 원본과 설정은 스킬 폴더에 복사하지 않는다. `source_rows`로 입력 위치를 추적할 수 있다. 실제 이름/학번은 출력하지 않는다.

서술 JSON은 계산을 읽은 다음 작성한다. 생략한 키는 짧은 자동 요약이 사용된다. `conclusion`은 정확히 3문단이다. 문장 속 수치는 analysis.json과 대조한다.

```json
{
  "distribution": "분포의 특징을 실제 수치에 맞춰 한두 문장으로 작성합니다.",
  "variability": "변동성이 큰 문항의 해석을 짧게 작성합니다.",
  "difficulty": "사용한 난이도 지표의 한계 내에서 해석합니다.",
  "discrimination": "변별력 지표를 해석하되 배점과 표본의 영향을 고려합니다.",
  "evaluation": ["데이터가 뒷받침하는 종합 평가", "구체적인 개선 제안"],
  "conclusion": ["전체 결과", "주요 문항과 개선 방향", "향후 활용 방안"]
}
```

의존성은 `scripts/requirements.txt`. 한글 글꼴을 자동으로 찾지 못하면 설치된 글꼴 경로를 `--font-path`로, Word 글꼴명은 `--doc-font`로 전달한다. 설치된 글꼴을 찾을 수 없으면 한글 그래프를 그대로 납품하지 않는다. 그림은 300dpi로 출력한다. 문항이 40개를 넘으면 여러 그림으로 분할하며 DOCX도 길어질 수 있다.
