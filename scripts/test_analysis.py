#!/usr/bin/env python3
"""Synthetic behavioral tests. No student data or persistent output."""
import copy
import csv
import json
import math
import tempfile
import unittest
from pathlib import Path
from analyze_exam import analyze, read_tables


class ExamTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self.rows = [
            ['id', '반', '1', '2', '3', '총점'],
            ['s1', 'A', 0, 0, 0, 0],
            ['s2', 'A', 1, 1, 2, 4],
            ['s3', 'A', 2, 2, 3, 7],
            ['s4', 'B', 2, 3, 4, 9],
            ['s5', 'B', '', '', '', 0],
            ['s6', 'B', 1, '', 1, 2],
            ['s7', '', 1, 2, 3, 6],
            ['s8', 'C', '', '', '', 0]]
        self.config = dict(title='가상 시험', sources=[dict(path='sample.csv', rows=list(range(2, 10)),
                           item_columns=['C', 'D', 'E'], class_column='B', total_column='F', id_column='A')],
                           items=[dict(label='1', max=2), dict(label='2', max=3), dict(label='3', max=5)],
                           total_max=10, max_source='가상 공식 배점')
        self.write()

    def tearDown(self):
        self.tmp.cleanup()

    def write(self):
        with (self.root / 'sample.csv').open('w', newline='', encoding='utf-8-sig') as f:
            csv.writer(f).writerows(self.rows)

    def run_analysis(self):
        return analyze(self.config, self.root)

    def test_missing_and_real_zero(self):
        a = self.run_analysis()
        self.assertEqual(a['counts']['complete_n'], 5)
        self.assertEqual(a['counts']['all_missing_n'], 2)
        self.assertEqual(a['counts']['partial_missing_n'], 1)
        self.assertEqual(a['counts']['zero_total_n'], 1)
        self.assertAlmostEqual(a['summary']['mean'], 5.2)
        self.assertAlmostEqual(a['summary']['sd'], math.sqrt(11.7))
        self.assertEqual(a['counts']['total_checked_n'], 5)
        self.assertEqual(a['items'][0]['positive_pct'], 80)
        self.assertEqual(a['items'][0]['earned_pct'], 60)
        self.assertEqual(a['items'][0]['normalized_delta'], .75)
        self.assertNotIn('s1', json.dumps(a))
        self.assertTrue(a['has_classes'])
        self.assertEqual(next(c for c in a['classes'] if c['label'] == 'C')['summary']['n'], 0)
        self.assertEqual(a['counts']['unclassified_n'], 1)

    def test_no_classes_or_maxima(self):
        self.config['sources'][0].pop('class_column')
        self.config.pop('total_max')
        self.config.pop('max_source')
        for item in self.config['items']:
            item.pop('max')
        a = self.run_analysis()
        self.assertFalse(a['has_classes'])
        self.assertFalse(a['normalized'])
        self.assertIsNone(a['total_max'])
        self.assertTrue(all(v['earned_pct'] is None for v in a['items']))

    def test_partial_maxima_do_not_mix_units(self):
        self.config['items'][1].pop('max')
        a = self.run_analysis()
        self.assertFalse(a['normalized'])
        self.assertIsNone(a['items'][1]['earned_pct'])
        self.assertEqual(a['items'][0]['earned_pct'], 60)

    def test_mismatched_total_stops(self):
        self.rows[1][-1] = 1
        self.write()
        with self.assertRaisesRegex(ValueError, '불일치'):
            self.run_analysis()

    def test_duplicates_stop(self):
        self.rows[2][0] = 's1'
        self.write()
        with self.assertRaisesRegex(ValueError, '중복 학생 ID'):
            self.run_analysis()

    def test_invalid_score_stops(self):
        for value in ['oops', '-1', 'nan', 'inf']:
            with self.subTest(value=value):
                self.rows[1][2] = value
                self.write()
                with self.assertRaises(ValueError):
                    self.run_analysis()

    def test_maximum_checks(self):
        self.rows[1][2] = 3
        self.write()
        with self.assertRaisesRegex(ValueError, '배점을 넘'):
            self.run_analysis()
        self.rows[1][2] = 0
        self.write()
        self.config['total_max'] = 11
        with self.assertRaisesRegex(ValueError, '배점 합'):
            self.run_analysis()

    def test_constant_zero_and_ties(self):
        for row in self.rows[1:]:
            row[2:] = [0, 0, 0, 0]
        self.write()
        a = self.run_analysis()
        self.assertEqual(a['counts']['complete_n'], 8)
        self.assertEqual(a['counts']['zero_total_n'], 8)
        self.assertIsNone(a['items'][0]['cv'])
        self.assertIsNone(a['items'][0]['delta'])
        self.assertIsNone(a['items'][0]['item_rest_r'])
        self.assertIn('중복', a['group']['unavailable_reason'])

    def test_boundary_ties_expanded(self):
        self.rows = self.rows[:1] + [[f'x{i}', 'A', v, 0, 0, v] for i, v in enumerate([0, 0, 0, 1, 1, 2, 2, 2])]
        self.write()
        a = self.run_analysis()
        self.assertEqual((a['group']['low_n'], a['group']['high_n']), (3, 3))
        self.assertEqual(a['items'][0]['delta'], 2)

    def test_single_student(self):
        self.config['sources'][0]['rows'] = [2]
        a = self.run_analysis()
        self.assertEqual(a['summary']['mean'], 0)
        self.assertIsNone(a['summary']['sd'])
        self.assertIsNone(a['items'][0]['sd'])
        self.assertIsNone(a['items'][0]['delta'])

    def test_explicit_exclusion_and_known_class(self):
        self.config['sources'][0].pop('class_column')
        self.config['sources'][0]['class_label'] = '확인된 반'
        self.config['sources'][0]['exclude_rows'] = {'2': '가상 등록 취소'}
        a = self.run_analysis()
        self.assertEqual(a['counts']['explicit_excluded_n'], 1)
        self.assertEqual(a['summary']['n'], 4)
        self.assertTrue(a['has_classes'])
        self.assertEqual(len(a['classes']), 1)

    def test_xlsx_multisheet_and_uncached_formula(self):
        from openpyxl import Workbook
        b = Workbook()
        for i, label in enumerate(['A', 'B']):
            sheet = b.active if i == 0 else b.create_sheet()
            sheet.title = label
            sheet.append(self.rows[0])
            for row in self.rows[1:5]:
                if row[1] == label:
                    sheet.append(row)
        b.save(self.root / 'sample.xlsx')
        src = copy.deepcopy(self.config['sources'][0])
        src.update(path='sample.xlsx', sheet='A', rows=[2, 3, 4])
        src2 = copy.deepcopy(src)
        src2.update(sheet='B', rows=[2])
        self.config['sources'] = [src, src2]
        a = self.run_analysis()
        self.assertEqual(a['summary']['n'], 4)
        self.assertEqual(a['summary']['mean'], 5)
        b['A']['C2'] = '=1+1'
        b.save(self.root / 'sample.xlsx')
        with self.assertRaisesRegex(ValueError, 'FORMULA_NO_CACHE'):
            self.run_analysis()

    def test_no_complete_cases(self):
        self.config['sources'][0]['rows'] = [6, 7]
        with self.assertRaisesRegex(ValueError, '완전한 문항'):
            self.run_analysis()


if __name__ == '__main__':
    unittest.main(verbosity=2)
