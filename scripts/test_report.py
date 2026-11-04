#!/usr/bin/env python3
"""Synthetic report integration checks; supply --out to retain results for visual QA."""
import argparse
import copy
import csv
import json
import tempfile
from pathlib import Path
from zipfile import ZipFile
import numpy as np
from docx import Document
from analyze_exam import analyze
from build_report import build


def run(root):
    root.mkdir(parents=True, exist_ok=True)
    maxima = [1, 2, 3] * 8
    random = np.random.default_rng(134)
    x = np.column_stack([random.integers(0, m + 1, size=32) for m in maxima])
    rows = [['id', '반'] + [str(i + 1) for i in range(24)] + ['총점']]
    rows += [[f's{i}', 'A반' if i < 16 else 'B반'] + list(map(int, row)) + [int(sum(row))] for i, row in enumerate(x)]
    rows += [['s32', 'C반'] + [''] * 24 + [0]]
    with (root / 'synthetic.csv').open('w', newline='', encoding='utf-8-sig') as f:
        csv.writer(f).writerows(rows)
    from openpyxl.utils import get_column_letter
    config = dict(title='가상 시험 검증', sources=[dict(path='synthetic.csv', first_row=2, last_row=34,
                  item_columns=[get_column_letter(i) for i in range(3, 27)], id_column='A', class_column='B', total_column='AA')],
                  items=[dict(label=str(i + 1), max=m) for i, m in enumerate(maxima)], max_source='가상 공식 배점표')
    a = analyze(config, root)
    narrative = {'evaluation': ['이 문서는 가상 자료로 생성 기능을 검증하기 위한 보고서입니다.'],
                 'conclusion': [f"가상 시험의 분석 학생은 {a['summary']['n']}명이며 평균은 {a['summary']['mean']:.2f}점이었습니다.",
                                '반별 문항 평균 득점률과 변별력의 계산 및 표시를 확인했습니다.',
                                '이 자료는 실제 학생의 시험 결과가 아닙니다.']}
    path, charts = build(a, narrative, root / 'classes')
    assert len(charts['heatmap']) == 1
    assert a['normalized'] and a['has_classes']
    assert len(Document(path).inline_shapes) == 5
    with ZipFile(root / 'classes' / '성적분석_그래프.zip') as z:
        assert len(z.namelist()) == 5
        assert any('히트맵' in name for name in z.namelist())
    # Constant scores must not crash KDE/violin or present undefined CV/delta as zero.
    for row in rows[1:]:
        row[2:] = [0] * 25
    with (root / 'synthetic.csv').open('w', newline='', encoding='utf-8-sig') as f:
        csv.writer(f).writerows(rows)
    unknown = copy.deepcopy(config)
    unknown['sources'][0].pop('class_column')
    for item in unknown['items']:
        item.pop('max')
    a = analyze(unknown, root)
    narrative['conclusion'] = ['가상 0점 시험입니다.', '모든 문항 점수가 같아 상·하위 변별력을 산출할 수 없습니다.', '이 자료는 실제 학생의 시험 결과가 아닙니다.']
    path, charts = build(a, narrative, root / 'constant')
    assert len(charts['heatmap']) == 0
    assert all(item['delta'] is None and item['cv'] is None for item in a['items'])
    assert len(Document(path).inline_shapes) == 4
    with ZipFile(root / 'constant' / '성적분석_그래프.zip') as z:
        assert len(z.namelist()) == 4
        assert not any('히트맵' in name for name in z.namelist())
    print(json.dumps({'class_report': str(root / 'classes' / '성적분석보고서.docx'),
                      'constant_report': str(path), 'result': 'passed'}, ensure_ascii=False))


if __name__ == '__main__':
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument('--out', type=Path)
    args = p.parse_args()
    if args.out:
        run(args.out)
    else:
        with tempfile.TemporaryDirectory() as directory:
            run(Path(directory))
