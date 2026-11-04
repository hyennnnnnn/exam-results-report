#!/usr/bin/env python3
"""Create concise Korean exam DOCX and 300-dpi chart ZIP from validated analysis."""
import argparse
import json
import math
import os
import tempfile
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
os.environ.setdefault('MPLCONFIGDIR', str(Path(tempfile.gettempdir()) / 'exam-report-mpl'))
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib import font_manager
from scipy.stats import gaussian_kde
from docx import Document
from docx.shared import Inches, Cm, Pt, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH

COLORS = ['#347BA2', '#CD7A43', '#4D8C7F', '#7A73A2', '#9A6676', '#727C44']


def fmt(value, digits=2):
    return '미산출' if value is None else f'{value:.{digits}f}'


def configure_font(path=None):
    if path:
        font_manager.fontManager.addfont(path)
        family = font_manager.FontProperties(fname=path).get_name()
    else:
        names = {f.name for f in font_manager.fontManager.ttflist}
        family = next((f for f in ['AppleGothic', 'Apple SD Gothic Neo', 'Malgun Gothic', 'NanumGothic', 'Noto Sans CJK KR'] if f in names), None)
        if family is None:
            raise ValueError('한글 글꼴이 없습니다. --font-path로 설치된 한글 글꼴을 지정하세요.')
    plt.rcParams.update({'font.family': family, 'font.size': 10, 'axes.unicode_minus': False,
                         'axes.spines.top': False, 'axes.spines.right': False, 'savefig.facecolor': 'white'})
    return family


def chunk(values, size):
    return [values[i:i + size] for i in range(0, len(values), size)]


def class_name(value):
    return '미분류' if value is None else str(value)


def make_figures(a, directory):
    directory.mkdir(parents=True)
    t, x = np.array(a['scores']), np.array(a['x'])
    items = a['items']
    files = dict(distribution=[], variability=[], profiles=[], difficulty=[], heatmap=[])
    counter = 0
    def save(fig, category, title):
        nonlocal counter
        counter += 1
        path = directory / f'{counter:02d}_{title}.png'
        fig.savefig(path, dpi=300, bbox_inches='tight', pad_inches=.16)
        plt.close(fig)
        files[category].append(path)
    def grid(ax):
        ax.grid(axis='y', color='#E2E6EB', lw=.6)
        ax.set_axisbelow(True)
    def item_ticks(ax, subset, capacity=18):
        step = max(1, math.ceil(len(subset) / capacity))
        positions = list(range(0, len(subset), step))
        ax.set_xticks(positions, [subset[i]['label'] for i in positions])
        ax.tick_params(axis='x', labelsize=8, rotation=45 if any(len(z['label']) > 3 for z in subset) else 0)

    def values_ax(ax, subset, key, title, ylabel):
        vals = [np.nan if z[key] is None else z[key] for z in subset]
        ax.bar(range(len(subset)), vals, color=COLORS[0])
        ax.set(title=title, ylabel=ylabel, xticks=range(len(subset)), xticklabels=[z['label'] for z in subset])
        item_ticks(ax, subset, 18 if ax.get_subplotspec().get_gridspec().ncols > 1 else 35)
        if any(z[key] is None for z in subset):
            ax.text(.99, .98, '빈 막대: 미산출', transform=ax.transAxes, ha='right', va='top', fontsize=8)
        grid(ax)
    groups = [(c['label'], t[np.array([v == c['label'] for v in a['row_classes']])]) for c in a['classes']] if a['has_classes'] else [('전체', t)]
    upper = a['total_max'] or max(1, float(t.max()) * 1.08)
    bins = np.linspace(0, upper, min(12, max(5, math.ceil(math.sqrt(len(t))))) + 1)
    for batch_index, batch in enumerate(chunk(groups, 6), 1):
        fig, axes = plt.subplots(2, 2, figsize=(8.6, 5), layout='constrained')
        for i, (name, vals) in enumerate(batch):
            label = f'{class_name(name)} (n={len(vals)})'
            color = COLORS[i % len(COLORS)]
            if len(vals) >= 3 and np.ptp(vals) > 0:
                xx = np.linspace(0, upper, 300)
                axes[0, 0].plot(xx, gaussian_kde(vals)(xx), color=color, label=label)
                violin = axes[1, 0].violinplot(vals, positions=[i + 1], showmedians=True)
                for body in violin['bodies']:
                    body.set_facecolor(color)
            else:
                axes[0, 0].plot(vals, np.zeros(len(vals)), '|', color=color, label=label + ' KDE 미산출')
                axes[1, 0].scatter(np.full(len(vals), i + 1), vals, color=color, s=15)
                axes[1, 0].text(i + 1, upper * .96, '밀도 미산출', ha='center', fontsize=7)
            if len(vals):
                axes[0, 1].hist(vals, bins=bins, color=color, alpha=.5, label=label)
                axes[1, 1].boxplot(vals, positions=[i + 1], widths=.5, patch_artist=True,
                                   boxprops={'facecolor': color, 'alpha': .4}, medianprops={'color': '#222222'})
        for ax, title in zip(axes.flat, ['밀도곡선 KDE', '히스토그램', '바이올린 그림', '상자그림']):
            ax.set_title(title)
            grid(ax)
        for ax in axes[0]:
            ax.set(xlabel='총점', xlim=(0, upper))
            ax.legend(fontsize=7, frameon=False)
        axes[0, 0].set_ylabel('밀도')
        axes[0, 1].set_ylabel('학생 수')
        for ax in axes[1]:
            ax.set(ylabel='총점', ylim=(0, upper * 1.02), xticks=range(1, len(batch) + 1),
                   xticklabels=[f'{class_name(name)}\n(n={len(vals)})' for name, vals in batch])
        save(fig, 'distribution', f'총점분포_{batch_index}')
    for start in range(0, len(items), 40):
        subset = items[start:start + 40]
        suffix = f'{start + 1}-{start + len(subset)}'
        fig, axes = plt.subplots(2, 2, figsize=(8.6, 4.8), layout='constrained')
        for ax, key, title, unit in zip(axes.flat, ['mean', 'sd', 'cv', 'range'], ['문항 평균', '표준편차', '변동계수', '관측 점수 범위'], ['점수', '점수', 'CV (%)', '최댓값 - 최솟값']):
            values_ax(ax, subset, key, title, unit)
        save(fig, 'variability', f'문항변동성_{suffix}')
        for batch_index, batch in enumerate(chunk(groups, 6), 1):
            fig, axes = plt.subplots(1, 2, figsize=(8.6, 3.5), layout='constrained')
            for i, (name, vals) in enumerate(batch):
                matrix = x[np.array([v == name for v in a['row_classes']])] if a['has_classes'] else x
                for ax, field, title in zip(axes, ['mean', 'sd'], ['문항별 평균 점수', '문항별 표준편차']):
                    profile = (matrix.mean(0) if field == 'mean' else matrix.std(0, ddof=1)) if len(matrix) >= (1 if field == 'mean' else 2) else np.full(x.shape[1], np.nan)
                    ax.plot(range(len(subset)), profile[start:start + len(subset)], marker='o', markersize=2.5,
                            color=COLORS[i], label=f'{class_name(name)} (n={len(matrix)})')
                    ax.set(title=title, ylabel='점수', xlabel='문항', xticks=range(len(subset)), xticklabels=[z['label'] for z in subset])
                    item_ticks(ax, subset)
                    grid(ax)
            for ax in axes:
                ax.legend(fontsize=7, frameon=False)
                ax.set_ylim(bottom=0)
            save(fig, 'profiles', f'문항평균과편차_{suffix}_{batch_index}')
        fig, axes = plt.subplots(2, 1, figsize=(8.6, 4.8), layout='constrained')
        values_ax(axes[0], subset, 'earned_pct' if a['normalized'] else 'positive_pct',
                  '평균 득점률' if a['normalized'] else '점수 획득률', '비율 (%)')
        axes[0].set_ylim(0, 108)
        values_ax(axes[1], subset, 'normalized_delta' if a['normalized'] else 'delta',
                  '상·하위 평균 차이 ÷ 배점' if a['normalized'] else '상·하위 평균 점수 차이',
                  '비율' if a['normalized'] else '점수 차이')
        axes[1].axhline(0, color='#666666', lw=.7)
        if a['group']['unavailable_reason']:
            axes[1].text(.5, .5, a['group']['unavailable_reason'], transform=axes[1].transAxes, ha='center')
        save(fig, 'difficulty', f'난이도와변별력_{suffix}')
        if a['has_classes']:
            for batch_index, batch in enumerate(chunk(a['classes'], 12), 1):
                heat = np.array([[c['items'][j]['mean'] if c['items'] else np.nan for j in range(start, start + len(subset))] for c in batch])
                if a['normalized']:
                    heat = heat / np.array([z['max_score'] for z in subset]) * 100
                fig, ax = plt.subplots(figsize=(8.6, max(2.3, .38 * len(batch) + 1.4)), layout='constrained')
                cmap = plt.get_cmap('Blues').copy()
                cmap.set_bad('#DDDDDD')
                vmax = 100 if a['normalized'] else max(1, max((v['mean'] for c in a['classes'] for v in c['items']), default=1))
                im = ax.imshow(np.ma.masked_invalid(heat), aspect='auto', cmap=cmap, vmin=0, vmax=vmax)
                ax.set(title='반별 문항 평균 득점률' if a['normalized'] else '반별 문항 평균 점수',
                       xticks=range(len(subset)), xticklabels=[z['label'] for z in subset], xlabel='문항',
                       yticks=range(len(batch)), yticklabels=[f"{class_name(c['label'])} (n={c['summary']['n']})" for c in batch])
                item_ticks(ax, subset, 35)
                if len(subset) <= 20:
                    for r in range(len(batch)):
                        for c in range(len(subset)):
                            val = heat[r, c]
                            ax.text(c, r, '—' if np.isnan(val) else f'{val:.1f}', ha='center', va='center', fontsize=7,
                                    color='white' if val > .65 * vmax else '#222222')
                fig.colorbar(im, ax=ax, label='득점률 (%)' if a['normalized'] else '평균 점수')
                save(fig, 'heatmap', f'반별히트맵_{suffix}_{batch_index}')
    return files


def build(a, narrative, out_dir, font_path=None, doc_font=None):
    if a.get('validation', {}).get('status') != 'passed':
        raise ValueError('계산 검증을 통과한 분석 결과가 필요합니다.')
    conclusion = narrative.get('conclusion')
    if not isinstance(conclusion, list) or len(conclusion) != 3 or any(not isinstance(v, str) or not v.strip() for v in conclusion):
        raise ValueError('계산 결과를 검토한 결론 3문단을 narrative에 작성하세요.')
    if out_dir.exists() and any(out_dir.iterdir()):
        raise ValueError('결과 폴더가 비어 있지 않습니다. 새 폴더를 사용하세요.')
    out_dir.mkdir(parents=True, exist_ok=True)
    family = configure_font(font_path)
    figures = make_figures(a, out_dir / '그래프')
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.top_margin = sec.bottom_margin = Cm(2.2)
    sec.left_margin = sec.right_margin = Cm(2.4)
    family = doc_font or ('Apple SD Gothic Neo' if family.startswith('Apple') else family)
    for name, size in [('Normal', 11), ('Title', 24), ('Heading 1', 16), ('Heading 2', 12), ('Caption', 9), ('Footer', 9), ('List Bullet', 11)]:
        style = doc.styles[name]
        style.font.name = family
        style.font.size = Pt(size)
        style.font.color.rgb = RGBColor(0, 0, 0)
        style.font.bold = name in ('Title', 'Heading 1', 'Heading 2')
        style.font.italic = False
        fonts = style.element.get_or_add_rPr().get_or_add_rFonts()
        for key in ['ascii', 'hAnsi', 'eastAsia', 'cs']:
            fonts.set(qn('w:' + key), family)
        for key in list(fonts.attrib):
            if 'Theme' in key:
                del fonts.attrib[key]
        pf = style.paragraph_format
        pf.line_spacing = 1.15
        pf.space_after = Pt(6)
        if name.startswith('Heading'):
            pf.keep_with_next = True
            pf.space_before = Pt(10)
    for border in list(doc.styles.element.iter(qn('w:pBdr'))):
        border.getparent().remove(border)
    doc.styles['Title'].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer = sec.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    field = OxmlElement('w:fldSimple')
    field.set(qn('w:instr'), 'PAGE')
    footer._p.append(field)
    doc.core_properties.title = a['title'] + ' 성적 분석 보고서'
    doc.core_properties.author = ''
    doc.core_properties.last_modified_by = ''
    def p(text, style='Normal'):
        par = doc.add_paragraph(text, style)
        par.paragraph_format.widow_control = True
        return par
    def h(text, page=False):
        par = p(text, 'Heading 1')
        if page:
            par.paragraph_format.page_break_before = True
    def bullet(text):
        return p(text, 'List Bullet')
    def note(text):
        return p(text, 'Caption')
    def images(category, caption):
        for path in figures[category]:
            par = doc.add_paragraph()
            par.paragraph_format.keep_with_next = True
            par.add_run().add_picture(str(path), width=Inches(6.2))
            note(caption)
    s, c, its, g = a['summary'], a['counts'], a['items'], a['group']
    def ranked(key, reverse=True, size=3):
        return sorted([z for z in its if z[key] is not None], key=lambda z: z[key], reverse=reverse)[:size]
    p(a['title'] + '\n성적 분석 보고서', 'Title')
    h('분석 개요')
    bullet(f"전체 {c['roster_n']}명 중 {s['n']}명 분석 / {len(its)}개 문항 / 만점 {fmt(a['total_max'])}")
    bullet(f"제외: 전 문항 결측 {c['all_missing_n']}명, 일부 결측 {c['partial_missing_n']}명, 명시적 제외 {c['explicit_excluded_n']}명")
    h('1 전체 및 반별 점수 분포' if a['has_classes'] else '1 전체 점수 분포')
    bullet(f"평균 {fmt(s['mean'])}점 / 표준편차 {fmt(s['sd'])} / 중앙값 {fmt(s['median'])}점")
    bullet(f"범위 {fmt(s['min'])}~{fmt(s['max'])}점 / 중간 50% {fmt(s['q1'])}~{fmt(s['q3'])}점")
    for cls in a['classes']:
        cs = cls['summary']
        bullet(f"{class_name(cls['label'])}: {cs['n']}명, 평균 {fmt(cs['mean'])}점, 표준편차 {fmt(cs['sd'])}")
    if narrative.get('distribution'):
        p(narrative['distribution'])
    images('distribution', '총점 분포. 소표본·동일 점수 집단의 밀도곡선은 미산출이며 관측점을 표시했습니다.' if len(a['scores']) < 3 or any(cls['summary']['n'] < 3 or cls['summary']['min'] == cls['summary']['max'] for cls in a['classes']) or s['min'] == s['max'] else '총점의 밀도곡선, 히스토그램, 바이올린 그림, 상자그림.')
    note(f"기준: 실제 0점은 포함하며 빈칸은 0점으로 바꾸지 않았습니다. 완전 사례를 주 분석에 사용했습니다. 문항 합계와 기록 총점 {c['total_checked_n']}건을 대조했고 평균·표본 표준편차를 독립 재계산했습니다. 사분위수는 선형 보간입니다.")
    if c['total_checked_n'] < s['n']:
        note(f"기록 총점이 없어 대조하지 못한 분석 학생: {s['n'] - c['total_checked_n']}명.")
    if c['id_unchecked_n']:
        note(f"고유 ID 미지정 {c['id_unchecked_n']}행은 중복 학생 여부를 확인하지 못했습니다.")
    if a['has_classes'] and c['unclassified_n']:
        note(f"분석 학생 중 반 정보 미확인 {c['unclassified_n']}명은 미분류로 표시했습니다.")
    h('2 문항별 점수 변동성', True)
    p('평균에 비해 점수 편차가 큰 문항을 변동계수로 살펴보았습니다.')
    for z in ranked('cv'):
        bullet(f"{z['label']}번: CV {fmt(z['cv'], 1)}% / 평균 {fmt(z['mean'])}점 / 표준편차 {fmt(z['sd'])}")
    for z in ranked('cv', False, 2):
        bullet(f"변동이 작은 {z['label']}번: CV {fmt(z['cv'], 1)}%")
    if narrative.get('variability'):
        p(narrative['variability'])
    images('variability', '문항별 평균, 표본 표준편차, 변동계수와 관측 점수 범위.')
    note('CV=표본 표준편차÷평균×100입니다. 평균이 작으면 커질 수 있어 변별력으로 해석하지 않습니다. 평균 0 또는 표본 1명의 CV는 미산출입니다. 관측 최댓값은 배점이 아닙니다.')
    h('3 문항 난이도', True)
    difficulty_key = 'earned_pct' if a['normalized'] else 'positive_pct'
    difficulty_label = '평균 득점률' if a['normalized'] else '점수 획득률'
    p('공식 배점에 대한 문항 평균의 비율로 살펴보았습니다.' if a['normalized'] else '문항 배점이 모두 확인되지 않아, 0점을 초과한 학생의 비율인 점수 획득률을 사용했습니다.')
    if len({round(z[difficulty_key], 10) for z in its}) == 1:
        bullet(f"모든 문항의 {difficulty_label}이 {fmt(its[0][difficulty_key], 1)}%로 같았습니다.")
    else:
        for z in ranked(difficulty_key, False):
            bullet(f"낮은 {difficulty_label}: {z['label']}번 {fmt(z[difficulty_key], 1)}%")
        for z in ranked(difficulty_key, True, 2):
            bullet(f"높은 {difficulty_label}: {z['label']}번 {fmt(z[difficulty_key], 1)}%")
    if narrative.get('difficulty'):
        p(narrative['difficulty'])
    images('profiles', '문항별 평균과 표본 표준편차. 표본 1명인 집단의 표준편차는 미산출입니다.')
    note('부분점수가 포함될 수 있으므로 정답률·만점률·통과율과 구별합니다. 원점수 평균과 편차는 배점 크기의 영향을 받습니다.')
    if a['max_source']:
        note('배점 근거: ' + a['max_source'])
    if a['has_classes']:
        h('반별 문항 비교', True)
        p('같은 색상 척도에서 반별 문항 성취 수준을 비교했습니다. 표본 수와 반 구성의 차이를 함께 고려해야 합니다.')
        images('heatmap', '반×문항 히트맵. 공식 배점 기준 평균 득점률(%)입니다. 회색은 미산출입니다.' if a['normalized'] else '반×문항 히트맵. 평균 원점수이며 문항별 배점 차이로 문항 간 직접 비교에 한계가 있습니다. 회색은 미산출입니다.')
        note('단순 평균 차이만으로 통계적 유의성이나 수업의 인과적 효과를 판단하지 않았습니다.')
    h('4 문항 변별력', True)
    if g['unavailable_reason']:
        p('상·하위 변별력은 미산출했습니다: ' + g['unavailable_reason'] + '.')
    else:
        p(f"총점 상위 {g['high_n']}명({fmt(g['high_cut'])}점 이상)과 하위 {g['low_n']}명({fmt(g['low_cut'])}점 이하)의 문항 평균 차이를 비교했습니다.")
        key = 'normalized_delta' if a['normalized'] else 'delta'
        for z in ranked(key):
            bullet(f"{z['label']}번: {'배점 대비 차이' if a['normalized'] else '평균 차이'} {fmt(z[key])}{'' if a['normalized'] else '점'}")
        negative = [z['label'] for z in its if z[key] is not None and z[key] < 0]
        if negative:
            bullet('상위 집단 평균이 더 낮은 문항: ' + '·'.join(negative))
    if narrative.get('discrimination'):
        p(narrative['discrimination'])
    images('difficulty', f'문항별 {difficulty_label}와 상·하위 집단 평균 차이.')
    note('상·하위 27% 인원을 올림하고 경계 동점자는 모두 포함했습니다. 배점으로 나눈 차이를 제시했습니다.' if a['normalized'] else '상·하위 27% 인원을 올림하고 경계 동점자는 모두 포함했습니다. 원점수 차이는 배점의 영향을 받으므로 문항 간 크기 비교는 참고 지표입니다.')
    h('5 종합 평가', True)
    for text in narrative.get('evaluation', [f"전체 평균은 {fmt(s['mean'])}점이며, 중간 50%는 {fmt(s['q1'])}~{fmt(s['q3'])}점에 분포했습니다.", '문항별 결과와 수업 목표를 함께 검토하여 보충학습 및 출제 개선 방향을 정할 필요가 있습니다.']):
        bullet(text)
    for text in a.get('context_notes', []):
        note(text)
    h('결론')
    conclusion = narrative.get('conclusion')
    if not isinstance(conclusion, list) or len(conclusion) != 3 or any(not isinstance(v, str) or not v.strip() for v in conclusion):
        raise ValueError('계산 결과를 검토한 결론 3문단을 narrative에 작성하세요.')
    for text in conclusion:
        p(text)
    note('자료: ' + ', '.join(dict.fromkeys(Path(z['path']).name + ' / ' + z['sheet'] for z in a['sources'])))
    path = out_dir / '성적분석보고서.docx'
    doc.save(path)
    with ZipFile(out_dir / '성적분석_그래프.zip', 'w', ZIP_DEFLATED) as archive:
        for category in figures.values():
            for file in category:
                archive.write(file, file.name)
    return path, figures


def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument('--analysis', type=Path, required=True)
    p.add_argument('--narrative', type=Path, required=True)
    p.add_argument('--out-dir', type=Path, required=True)
    p.add_argument('--font-path')
    p.add_argument('--doc-font')
    args = p.parse_args()
    a = json.loads(args.analysis.read_text(encoding='utf-8'))
    narrative = json.loads(args.narrative.read_text(encoding='utf-8'))
    doc, figures = build(a, narrative, args.out_dir, args.font_path, args.doc_font)
    print(json.dumps(dict(docx=str(doc.resolve()), charts=sum(map(len, figures.values())), heatmaps=len(figures['heatmap'])), ensure_ascii=False))


if __name__ == '__main__':
    main()
