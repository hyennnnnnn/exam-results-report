#!/usr/bin/env python3
"""Read mapped exam scores without mutating inputs; emit aggregate audit + anonymous matrix."""
import argparse
import csv
import json
import math
import statistics
from pathlib import Path
import numpy as np


def column(value):
    if not isinstance(value, str) or not value.isascii() or not value.isalpha():
        raise ValueError(f'열은 Excel 문자여야 합니다: {value!r}')
    result = 0
    for char in value.upper():
        result = result * 26 + ord(char) - 64
    return result - 1


def read_tables(path, encoding='utf-8-sig', delimiter=','):
    suffix = path.suffix.lower()
    if suffix == '.xls':
        import xlrd
        book = xlrd.open_workbook(str(path))
        return {s.name: [[('#ERROR' if s.cell_type(r, c) == xlrd.XL_CELL_ERROR else
                           bool(s.cell_value(r, c)) if s.cell_type(r, c) == xlrd.XL_CELL_BOOLEAN else
                           s.cell_value(r, c)) for c in range(s.ncols)] for r in range(s.nrows)]
                for s in book.sheets()}
    if suffix == '.xlsx':
        from openpyxl import load_workbook
        book = load_workbook(path, data_only=True, read_only=True)
        result = {s.title: [list(row) for row in s.iter_rows(values_only=True)] for s in book}
        book.close()
        formulas = load_workbook(path, data_only=False, read_only=True)
        for sheet in formulas:
            for row in sheet.iter_rows():
                for cell in row:
                    if cell.data_type == 'f' and result[sheet.title][cell.row - 1][cell.column - 1] is None:
                        result[sheet.title][cell.row - 1][cell.column - 1] = '#FORMULA_NO_CACHE'
        formulas.close()
        return result
    if suffix == '.csv':
        with path.open(encoding=encoding, newline='') as stream:
            return {'CSV': list(csv.reader(stream, delimiter=delimiter))}
    raise ValueError('지원 형식은 XLS, XLSX, CSV입니다.')


def number(value, missing_tokens=()):
    if value is None or (isinstance(value, str) and (not value.strip() or value.strip() in missing_tokens)):
        return None
    if isinstance(value, bool):
        raise ValueError('참/거짓은 점수로 사용할 수 없습니다.')
    result = float(value)
    if not math.isfinite(result) or result < 0:
        raise ValueError(f'유효하지 않은 점수: {value!r}')
    return result


def describe(values):
    values = np.asarray(values, float)
    if not len(values):
        return dict(n=0, mean=None, sd=None, median=None, min=None, max=None, q1=None, q3=None)
    return dict(n=len(values), mean=float(values.mean()),
                sd=float(values.std(ddof=1)) if len(values) > 1 else None,
                median=float(np.median(values)), min=float(values.min()), max=float(values.max()),
                q1=float(np.quantile(values, .25)), q3=float(np.quantile(values, .75)))


def item_summary(x):
    return [dict(mean=float(v.mean()), sd=float(v.std(ddof=1)) if len(v) > 1 else None)
            for v in x.T] if len(x) else []


def analyze(config, base=Path('.')):
    items = config['items']
    labels = [str(i['label']) for i in items]
    if not items or len(set(labels)) != len(labels):
        raise ValueError('문항은 하나 이상이고 문항명이 고유해야 합니다.')
    maxima = [number(i.get('max')) for i in items]
    total_max = number(config.get('total_max'))
    if any(m == 0 for m in maxima) or total_max == 0:
        raise ValueError('공식 배점은 양수여야 합니다.')
    if (any(m is not None for m in maxima) or total_max is not None) and not config.get('max_source', '').strip():
        raise ValueError('공식 배점의 출처 max_source가 필요합니다.')
    normalized = all(m is not None for m in maxima)
    if normalized:
        if total_max is not None and not math.isclose(sum(maxima), total_max, abs_tol=1e-8, rel_tol=1e-9):
            raise ValueError('문항 배점 합과 공식 총점 만점이 다릅니다.')
        total_max = sum(maxima)
    if not config.get('sources'):
        raise ValueError('입력 자료가 없습니다.')
    matrix, classes, source_rows, exclusions = [], [], [], []
    seen_ids, seen_rows, known_classes, stored_totals = set(), set(), set(), []
    counts = dict(roster_n=0, complete_n=0, all_missing_n=0, partial_missing_n=0, explicit_excluded_n=0,
                  zero_item_n=0, zero_total_n=0, total_checked_n=0, id_unchecked_n=0,
                  unclassified_n=0)
    all_sources = []
    tokens = config.get('missing_tokens', [])
    for source in config['sources']:
        path = (base / source['path']).resolve()
        tables = read_tables(path, source.get('encoding', 'utf-8-sig'), source.get('delimiter', ','))
        sheet = source.get('sheet', 'CSV' if path.suffix.lower() == '.csv' else None)
        if sheet not in tables:
            raise ValueError(f'시트를 명시하거나 확인하세요: {sheet!r}')
        rows = tables[sheet]
        if 'rows' in source and ('first_row' in source or 'last_row' in source):
            raise ValueError('rows와 first_row/last_row를 함께 쓰지 않습니다.')
        selected = source.get('rows')
        if selected is None:
            selected = list(range(source['first_row'], source['last_row'] + 1))
        if not selected:
            raise ValueError('학생 행 선택이 비어 있습니다.')
        cols = [column(c) for c in source['item_columns']]
        if len(cols) != len(items) or len(set(cols)) != len(cols):
            raise ValueError('문항 열은 문항과 1:1로 대응해야 합니다.')
        if 'class_column' in source and 'class_label' in source:
            raise ValueError('반 열과 반 이름 중 하나만 지정합니다.')
        metadata_cols = [column(source[key]) for key in ['id_column', 'class_column', 'total_column'] if key in source]
        if len(set(metadata_cols + cols)) != len(metadata_cols + cols):
            raise ValueError('문항/총점/반/ID 열이 겹칩니다.')
        excluded_map = source.get('exclude_rows', {})
        if not set(map(int, excluded_map)).issubset(set(selected)):
            raise ValueError('제외 행은 선택한 학생 행에 속해야 합니다.')
        all_sources.append(dict(path=str(path), sheet=sheet, selected_rows=selected))
        for row_num in selected:
            if isinstance(row_num, bool) or not isinstance(row_num, int) or row_num < 1 or row_num > len(rows):
                raise ValueError(f'학생 행 범위 오류: {row_num!r}')
            key = (str(path), sheet, row_num)
            if key in seen_rows:
                raise ValueError('같은 학생 행이 중복 선택되었습니다.')
            seen_rows.add(key)
            counts['roster_n'] += 1
            locator = dict(source=len(all_sources) - 1, row=row_num)
            row = rows[row_num - 1]
            def get(col):
                index = column(col) if isinstance(col, str) else col
                return row[index] if index < len(row) else None
            if str(row_num) in excluded_map:
                reason = excluded_map[str(row_num)]
                if not isinstance(reason, str) or not reason.strip():
                    raise ValueError('명시적 제외에는 사유가 필요합니다.')
                exclusions.append(dict(**locator, reason=reason))
                counts['explicit_excluded_n'] += 1
                continue
            if source.get('id_column'):
                value = get(source['id_column'])
                if value is None or str(value).strip() == '':
                    raise ValueError(f'{sheet} {row_num}행 학생 ID가 비어 있습니다.')
                identity = str(int(value)) if isinstance(value, (int, float)) and not isinstance(value, bool) and float(value).is_integer() else str(value).strip()
                if identity in seen_ids:
                    raise ValueError('선택 자료에 중복 학생 ID가 있습니다. 원본을 확인하세요.')
                seen_ids.add(identity)
            else:
                counts['id_unchecked_n'] += 1
            cls = get(source['class_column']) if source.get('class_column') else source.get('class_label')
            cls = str(cls).strip() if cls is not None and str(cls).strip() else None
            if cls is not None:
                known_classes.add(cls)
            try:
                values = [number(get(c), tokens) for c in cols]
                recorded = number(get(source['total_column']), tokens) if source.get('total_column') else None
            except (ValueError, TypeError) as exc:
                raise ValueError(f'{sheet} {row_num}행: {exc}') from exc
            if recorded is not None:
                stored_totals.append(recorded)
            for value, maximum in zip(values, maxima):
                if value is not None and maximum is not None and value > maximum + 1e-8:
                    raise ValueError(f'{sheet} {row_num}행 점수가 공식 문항 배점을 넘습니다.')
            n_missing = sum(v is None for v in values)
            if n_missing:
                category = 'all_missing_n' if n_missing == len(items) else 'partial_missing_n'
                counts[category] += 1
                exclusions.append(dict(**locator, reason='전 문항 결측' if n_missing == len(items) else '일부 문항 결측'))
                continue
            total = math.fsum(values)
            if total_max is not None and total > total_max + 1e-8:
                raise ValueError(f'{sheet} {row_num}행 총점이 공식 만점을 넘습니다.')
            if recorded is not None:
                if not math.isclose(total, recorded, abs_tol=1e-8, rel_tol=1e-9):
                    raise ValueError(f'{sheet} {row_num}행 문항 합계와 기록된 총점 불일치: {total} / {recorded}')
                counts['total_checked_n'] += 1
            matrix.append(values)
            classes.append(cls)
            source_rows.append(locator)
    if not matrix:
        raise ValueError('완전한 문항 점수가 있는 학생이 없습니다. 결측 처리 기준을 검토하세요.')
    x = np.array(matrix, dtype=float)
    t = x.sum(axis=1)
    n, m = x.shape
    counts.update(complete_n=n, zero_item_n=int((x == 0).sum()), zero_total_n=int((t == 0).sum()),
                  unclassified_n=sum(c is None for c in classes))
    k = math.ceil(n * .27)
    low_cut, high_cut = float(np.sort(t)[k-1]), float(np.sort(t)[-k])
    low, high = t <= low_cut, t >= high_cut
    group_reason = '표본 4명 미만' if n < 4 else '경계 동점으로 상·하위 집단 중복' if np.any(low & high) else None
    group = dict(target_n=k, low_n=int(low.sum()), high_n=int(high.sum()), low_cut=low_cut,
                 high_cut=high_cut, unavailable_reason=group_reason)
    results, errors = [], []
    for j in range(m):
        v = x[:, j]
        mean = float(v.mean())
        sd = float(v.std(ddof=1)) if n > 1 else None
        delta = float(v[high].mean() - v[low].mean()) if group_reason is None else None
        rest = t - v
        correlation = float(np.corrcoef(v, rest)[0, 1]) if np.ptp(v) > 0 and np.ptp(rest) > 0 else None
        errors.append(abs(mean - statistics.mean(v.tolist())))
        if n > 1:
            errors.append(abs(sd - statistics.stdev(v.tolist())))
        if correlation is not None:
            errors.append(abs(correlation - statistics.correlation(v.tolist(), rest.tolist())))
        results.append(dict(label=labels[j], max_score=maxima[j], mean=mean, sd=sd,
                            cv=sd / mean * 100 if mean > 0 and sd is not None else None,
                            min=float(v.min()), max=float(v.max()), range=float(np.ptp(v)),
                            positive_n=int((v > 0).sum()), positive_pct=float((v > 0).mean() * 100),
                            earned_pct=mean / maxima[j] * 100 if maxima[j] is not None else None,
                            delta=delta, normalized_delta=delta / maxima[j] if delta is not None and maxima[j] is not None else None,
                            item_rest_r=correlation))
    mean_error = abs(sum(item['mean'] for item in results) - float(t.mean()))
    errors.append(mean_error)
    errors.append(abs(statistics.mean([math.fsum(v) for v in matrix]) - float(t.mean())))
    if max(errors, default=0) > 1e-8 * max(1, float(t.max())):
        raise ValueError('독립 재계산 검증에 실패했습니다.')
    class_stats = []
    if known_classes:
        group_labels = sorted(known_classes)
        if any(c is None for c in classes):
            group_labels.append(None)
        for label in group_labels:
            mask = np.array([c == label for c in classes])
            class_stats.append(dict(label=label, summary=describe(t[mask]), items=item_summary(x[mask])))
    return dict(title=config.get('title', '시험'), summary=describe(t), counts=counts, items=results,
                total_max=total_max, normalized=normalized, max_source=config.get('max_source'),
                has_classes=bool(known_classes), classes=class_stats, row_classes=classes,
                group=group, scores=t.tolist(), x=matrix, sources=all_sources, source_rows=source_rows,
                exclusions=exclusions, context_notes=config.get('context_notes', []),
                validation=dict(max_numeric_error=max(errors, default=0),
                                recorded_totals=describe(stored_totals), status='passed'))


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest='command', required=True)
    inspect = sub.add_parser('inspect')
    inspect.add_argument('file', type=Path)
    inspect.add_argument('--encoding', default='utf-8-sig')
    calculate = sub.add_parser('analyze')
    calculate.add_argument('config', type=Path)
    calculate.add_argument('--out', required=True, type=Path)
    args = parser.parse_args()
    if args.command == 'inspect':
        tables = read_tables(args.file, args.encoding)
        print(json.dumps({name: dict(rows=len(rows), cols=max(map(len, rows), default=0), preview=[row[:12] for row in rows[:8]])
                          for name, rows in tables.items()}, ensure_ascii=False, indent=2, default=str))
    else:
        config = json.loads(args.config.read_text(encoding='utf-8'))
        result = analyze(config, args.config.parent)
        args.out.parent.mkdir(parents=True, exist_ok=True)
        args.out.write_text(json.dumps(result, ensure_ascii=False, indent=2, allow_nan=False), encoding='utf-8')
        print(json.dumps(dict(summary=result['summary'], counts=result['counts'], has_classes=result['has_classes'],
                              normalized=result['normalized'], validation=result['validation']), ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
