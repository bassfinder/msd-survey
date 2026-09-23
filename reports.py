# -*- coding: utf-8 -*-
"""근골격계부담작업 유해요인조사 결과보고서 — Excel / PDF 생성

입력은 app.analyze(round) 결과(dict A) 하나. 화면·Excel·PDF가 같은 집계를 쓰므로 숫자가 어긋나지 않는다.
"""
import io, json, os
from datetime import datetime, timedelta
import survey_def as S

KST = timedelta(hours=9)
LEGAL = ('산업안전보건법 제39조(보건조치) 및 산업안전보건기준에 관한 규칙 제656조~제667조(근골격계부담작업으로 인한 건강장해의 예방). '
         '제657조에 따라 근골격계부담작업을 하는 근로자가 있는 경우 3년마다 유해요인조사를 실시하여야 하며, '
         '신설 사업장은 신설일부터 1년 이내에 최초 조사를 실시한다.')
METHOD = [
    ('근골격계부담작업 해당 여부', '고용노동부 고시 근골격계부담작업 제1~11호 해당 여부를 단위작업별로 확인'),
    ('유해요인조사표([별지 제1호서식])', '조사개요, 작업장 상황(작업설비·작업량·작업속도·업무변화), 작업조건(1단계 작업내용, 2단계 단위작업별 부담작업(호)·작업부하 A × 작업빈도 B, 3단계 유해요인 원인분석)'),
    ('근골격계질환 증상조사표([별지 제2호서식])', '대상 근로자 자기기입식 설문(모바일, 다국어) — 부위별 통증 기간·정도·빈도 조사'),
]


def jl(s, d=None):
    try:
        return json.loads(s) if s else (d if d is not None else {})
    except Exception:
        return d if d is not None else {}


def fmt_date(d):
    return d.strftime('%Y.%m.%d') if d else ''


def period(r):
    if not r:
        return ''
    return f'{fmt_date(r.start_date)} ~ {fmt_date(r.end_date)}'.strip(' ~')


def conclusion_text(A):
    """종합의견 — 회차에 직접 입력한 게 있으면 그걸, 없으면 집계로 자동 작성"""
    r = A['round']
    if r and r.conclusion:
        return [p for p in r.conclusion.split('\n') if p.strip()]
    n = A['n']
    out = []
    if not n:
        return ['증상조사 응답이 없어 결과를 산출하지 않았습니다.']
    pp = A['persons_p']
    out.append(f"증상조사 응답자 {n}명 중 증상호소자는 {A['persons']['complaint']}명({pp['complaint']}%), "
               f"관리대상자는 {A['persons']['manage']}명({pp['manage']}%), "
               f"통증호소자는 {A['persons']['pain']}명({pp['pain']}%)으로 나타났다.")
    top = sorted([p for p in A['parts'] if p['manage']], key=lambda p: -p['manage'])[:3]
    if top:
        out.append('신체부위별로는 ' + ', '.join(f"{p['label']}({p['manage_p']}%)" for p in top) +
                   ' 순으로 관리대상 기준 통증호소율이 높았다.')
    burden = [x for x in A['job_summary'] if x['is_burden']]
    if A['assessments']:
        out.append(f"조사대상 작업 {len(A['assessments'])}개 중 근골격계부담작업(제1~11호)에 해당하는 작업은 {len(burden)}개이다."
                   + (' 해당 작업: ' + ', '.join(f"{x['job'].name}({','.join(map(str, x['items']))}호)" for x in burden) if burden else ''))
    if A['ranked']:
        out.append('개선 우선순위는 ' + ', '.join(f"{x['rank']}순위 {x['job'].name}" for x in A['ranked'][:5]) +
                   ' 으로 선정하였으며, 작업환경 개선계획서에 따라 개선을 추진한다.')
    q1 = [x['job'].name for x in A['quad'].get(1, [])]
    q2 = [x['job'].name for x in A['quad'].get(2, [])]
    q3 = [x['job'].name for x in A['quad'].get(3, [])]
    if q1 or q2 or q3:
        out.append('유해요인조사표와 증상조사표를 종합한 결과, '
                   + (f"유해요인과 증상이 모두 확인된 즉시 개선 대상은 {', '.join(q1)}" if q1 else '즉시 개선 대상 작업은 없으며')
                   + (f", 증상은 적으나 유해요인이 높아 예방적 개선이 필요한 작업은 {', '.join(q2)}" if q2 else '')
                   + (f", 유해요인 점수에 비해 증상이 많아 재확인이 필요한 작업은 {', '.join(q3)}" if q3 else '') + '이다.')
    fits = [x for x in A['job_summary'] if x.get('match')]
    if fits:
        out.append('부담작업의 관련 신체부위와 실제 증상 부위가 일치하는 작업(' + ', '.join(
            f"{x['job'].name}: {'·'.join(S.label('part', p) for p in x['match'])}" for x in fits) + ')은 작업 관련성이 높은 것으로 판단된다.')
    out.append('통증호소자에 대해서는 보건관리자 상담 및 필요 시 의료기관 진료를 권고하고, 작업 전환·작업방법 개선 등 사후관리를 실시한다.')
    out.append('근골격계부담작업 종사자에게 유해요인, 징후와 증상, 올바른 작업자세 및 스트레칭 등을 알리고(안전보건규칙 제661조), '
               '개선 완료 후 효과를 평가하여 다음 조사에 반영한다.')
    return out


QORDER = (1, 2, 5, 3, 4, 0)


def summary_data(A):
    """1쪽 요약용 숫자·문장"""
    r = A['round']
    n = A['n']
    active_jobs = [j for j in A['job_summary'] if j['job'].active]
    burden = [j for j in A['job_summary'] if j['is_burden']]
    imps = A['improvements']
    fol = A['pain_list'] + A['manage_list']
    top_parts = sorted([p for p in A['parts'] if p['manage']], key=lambda p: -p['manage'])[:3]
    return {
        'overview': [
            ('조사', f"{r.title if r else ''} ({r.kind if r else ''}) · {period(r)}"),
            ('대상', f"작업 {len(active_jobs)}개 · 근로자 {A['target_workers'] or '-'}명 · 증상조사 응답 {n}명"
                    + (f" (응답률 {round(n * 100 / A['target_workers'], 1)}%)" if A['target_workers'] else '')),
            ('부담작업', f"유해요인조사표 작성 {len(A['assessments'])}개 작업 중 근골격계부담작업 해당 {len(burden)}개"
                       + (': ' + ', '.join(f"{j['job'].name}({','.join(map(str, j['items']))}호)" for j in burden) if burden else '')),
            ('증상', f"관리대상자 {A['persons']['manage']}명({A['persons_p']['manage']}%) · 통증호소자 {A['persons']['pain']}명({A['persons_p']['pain']}%)"
                    + (' · 주요 부위 ' + ', '.join(f"{p['label']} {p['manage_p']}%" for p in top_parts) if top_parts else '')),
            ('사후관리', f"관리대상자 이상 {len(fol)}명 중 조치 {sum(1 for x in fol if x.followup and x.followup != '미조치')}명 · 미조치 {sum(1 for x in fol if not x.followup or x.followup == '미조치')}명"),
            ('개선계획', f"{len(imps)}건 · 예산 {sum(m.cost or 0 for m in imps):,}천원 · 완료 {sum(1 for m in imps if m.status == '완료')}건"),
        ],
        'quad': [(S.QUADRANT[q][0], ', '.join(x['job'].name for x in A['quad'][q]) or '-', S.QUADRANT[q][1])
                 for q in QORDER if A['quad'].get(q)],
        'top': [[x['rank'], x['q_label'], f"{x['job'].dept} / {x['job'].name}", x['why']] for x in A['ranked'][:5]],
    }


def judge_rows(A):
    """5장 종합판정 표 데이터"""
    basis = [[x['q_label'], f"{x['job'].dept} / {x['job'].name}", ','.join(map(str, x['items'])) or '-',
              x['score'] or '-', x['n'], f"{x['manage']} ({x['rate']}%)", x['pain'], x['why']]
             for q in QORDER for x in A['quad'].get(q, [])]
    fit = []
    for x in A['job_summary']:
        if not (x['n'] or x['items']):
            continue
        cells = []
        for p in S.PARTS:
            sp = x['sym_parts'][p]
            hit = sp['manage'] + sp['pain']
            mark = ('●' if p in x['exp_parts'] else '')
            if hit:
                mark += f" {sp['manage']}/{sp['pain']}" + (' ◎' if p in x['match'] else (' △' if p in x['unexpected'] and x['assess'] else ''))
            cells.append(mark.strip())
        fit.append([f"{x['job'].dept} / {x['job'].name}", ','.join(map(str, x['items'])) or '-'] + cells + [x['q_label']])
    pers = [[p['r'].name, S.KOSHA_KO[p['r'].top_level], p['r'].job_label, ', '.join(S.label('part', q) for q in p['parts']),
             ', '.join(S.label('part', q) for q in p['work_parts']) or '-', ' · '.join(p['flags']) or '없음', p['judge'],
             p['r'].followup or '미조치'] for p in A['personal']]
    return basis, fit, pers


JUDGE_BASIS = [
    ('유해요인 높음', f"근골격계부담작업(제1~11호) 해당 + 단위작업 최고 총점수(작업부하 A × 작업빈도 B) {S.EXPOSURE_SCORE_HIGH}점 이상"),
    ('증상 많음', '해당 작업에 통증호소자가 있거나, 관리대상자 이상 비율이 사업장 전체 평균 이상 (편람: 부서별 증상호소율 비교)'),
    ('부위 부합성', '부담작업이 부담을 주는 신체부위(고시 체크리스트 기준)에 실제 관리대상자·통증호소자가 있으면 ◎ — 작업 관련성의 근거'),
    ('개인요인', '증상 부위와 겹치는 과거 상해·질병·취미·가사노동(증상조사표 I부)을 표시해 작업 외 요인을 함께 검토'),
    ('우선순위', '① 즉시 개선 → ② 예방적 개선 → ※ 조사표 미작성 → ③ 재확인 → ④ 유지 순, 같은 판정 안에서는 통증호소자·관리대상자 수, A×B, 노출 인원 순'),
]


def raw_rows(A):
    head = ['번호', '제출일시(KST)', '언어', '성명', '직원명단ID', '연령', '성별', '부서', '수행작업', '직장경력(년)', '직장경력(월)',
            '결혼', '현재작업', '현작업기간(년)', '현작업기간(월)', '1일근무(시간)', '휴식(분)', '휴식(회)', '이전작업',
            '여가·취미', '가사노동', '질병진단', '진단질병', '질병상태', '과거상해', '상해부위', '육체적부담', '통증유무']
    for p in S.PARTS:
        lb = S.label('part', p)
        head += [f'{lb}_판정', f'{lb}_좌우', f'{lb}_기간', f'{lb}_정도', f'{lb}_빈도', f'{lb}_지난1주', f'{lb}_통증으로인한일']
    head += ['개인판정', '선택 작업(작업부하A/작업빈도B)', '개선의견', '사후관리', '관리메모']
    rows = []
    for r in A['responses']:
        pdct = r.pain_dict
        lv = r.part_levels
        row = [r.id, (r.created_at + KST).strftime('%Y-%m-%d %H:%M') if r.created_at else '', S.LANG_NAMES.get(r.lang, r.lang),
               r.name, r.employee_id, r.age, S.label('sex', r.sex), r.dept, r.job_label, r.career_y, r.career_m,
               S.label('married', r.married), r.cur_task_ko or r.cur_task, r.cur_y, r.cur_m, r.work_hours, r.rest_min, r.rest_times,
               r.prev_task_ko or r.prev_task,
               ', '.join(S.label('hobby', h) for h in jl(r.hobbies, [])), S.label('housework', r.housework),
               {'Y': '예', 'N': '아니오'}.get(r.disease_yn, ''), ', '.join(S.label('disease', d) for d in jl(r.diseases, [])),
               S.label('dstatus', r.disease_status) if r.disease_yn == 'Y' else '',
               {'Y': '예', 'N': '아니오'}.get(r.injury_yn, ''),
               ', '.join(S.label('injury', d) for d in jl(r.injuries, [])),
               S.label('burden', r.burden), {'Y': '예', 'N': '아니오'}.get(r.has_pain, '')]
        for p in S.PARTS:
            d = pdct.get(p)
            if d:
                cq = ', '.join(S.label('cq', c) for c in d.get('conseq') or [])
                if d.get('other'):
                    cq += f" ({d.get('other_ko') or d.get('other')})"
                row += [S.LEVEL_KO[lv.get(p)], S.label('side', d.get('side')) if d.get('side') else '', S.label('dur', d.get('duration')),
                        S.label('int', d.get('intensity')), S.label('freq', d.get('frequency')),
                        {'Y': '예', 'N': '아니오'}.get(d.get('lastweek'), ''), cq]
            else:
                row += [''] * 7
        row += [S.LEVEL_KO[r.top_level], ', '.join(f"{u.get('name')}({u.get('load') or '-'}/{u.get('freq') or '-'})" for u in jl(r.units, [])),
                r.opinion_ko or r.opinion, r.followup or '', r.admin_note or '']
        rows.append(row)
    return head, rows


KOSHA_HEAD = ['#', '성명', '연령', '성별(남1/여2)', '현 직장경력', '부서', '작업', '현재 작업내용', '작업기간',
              '목', '어깨', '팔', '손', '허리', '다리', '전체 결과']
KOSHA_NOTE = "※ KOSHA Code H-30-2008 양식 활용. 본 결과는 자기기입식 설문에 따른 분류로, 의학적 진단이나 법률적 판단의 근거로 사용할 수 없음."


def _ym(y, m):
    if y is None and m is None:
        return ''
    if not m:
        return f'{y or 0}년'
    return f'{y}년{m}개월' if y else f'{m}개월'


def kosha_rows(A):
    """2023년 위탁조사 결과표와 같은 형식: 부위별 정상/관리대상자/통증호소자 + 전체 결과"""
    rows = []
    for i, r in enumerate(A['responses'], 1):
        lv = r.part_levels
        rows.append([i, r.name, f'{r.age}세' if r.age else '', {'M': 1, 'F': 2}.get(r.sex, ''), _ym(r.career_y, r.career_m),
                     r.dept, r.job_label, r.cur_task_ko or r.cur_task or '', _ym(r.cur_y, r.cur_m)] +
                    [S.KOSHA_KO[lv.get(p)] for p in S.PARTS] + [S.KOSHA_KO[r.top_level]])
    return rows


# ── 유해요인조사표([별지 제1호서식]) 공통 ─────────────────────────────
def unit_items(a):
    """단위작업명 → 부담작업(호). 옛 데이터(체크리스트 별도 저장)도 지원"""
    m = {u.get('unit'): [int(i) for i in u.get('items', [])] for u in jl(a.checklist, [])}
    for t in a.task_list:
        if t.get('items') is not None:
            m[t.get('name')] = [int(i) for i in t['items']]
    return m


def unit_workers(a):
    return {t.get('name'): t.get('workers') for t in a.task_list} if a else {}


def ho(items):
    return ', '.join(str(i) for i in items) if items else '-'


def f1_sections(a, js, r):
    """[별지 제1호서식] 항목을 섹션별로 — Excel·PDF 공용"""
    job = js['job']
    st = jl(a.site_status)
    um = unit_items(a)
    tasks = a.task_list
    hz = jl(a.hazards, [])
    ergo = jl(a.ergo, []) if getattr(a, 'ergo', None) else []
    site = []
    for k, lab in S.SITE_ITEMS:
        v = st.get(k, {})
        val = v.get('v') or '변화 없음'
        site.append((lab, val + (f"  ({'언제부터: ' if val in ('변화 있음', '줄음', '늘어남') else ''}{v.get('since')})" if v.get('since') else '')))
    t_rows = []
    for t in tasks:
        ld, fq = int(t.get('load') or 0), int(t.get('freq') or 0)
        t_rows.append([t.get('name'), ho(um.get(t.get('name'), [])),
                       f"{ld} {dict(S.LOAD_SCORE).get(ld, '')}" if ld else '',
                       f"{fq} {dict(S.FREQ_SCORE).get(fq, '')}" if fq else '', ld * fq or ''])
    h_rows = [[h.get('task'), ho(h.get('items') or um.get(h.get('task'), [])), h.get('type'), h.get('cause'),
               h.get('note') or h.get('parts') or ''] for h in hz]
    e_rows = [[e.get('unit'), e.get('tool'), e.get('result'), e.get('judge')] for e in ergo]
    summary = getattr(a, 'summary', None)
    if not summary:
        types = sorted({h.get('type') for h in hz if h.get('type')})
        summary = (f"근골격계부담작업: {', '.join(f'제{i}호' for i in a.burden_items) or '비해당'}. "
                   f"단위작업 최고 총점수(A×B) {a.max_score}점."
                   + (f" 주요 유해요인: {', '.join(types)}." if types else ''))
    symptom = (f"응답 {js['n']}명 — 관리대상자 {js['manage'] - js['pain']}명, 통증호소자 {js['pain']}명"
               if js['n'] else '해당 작업 응답 없음')
    return {
        'overview': [('조사일시 / 조사자', f'{fmt_date(a.survey_date)} / {a.investigator or ""}'),
                     ('부서명', job.dept), ('작업공정명', job.process or ''), ('작업명', a.job_title or job.name)],
        'site': site,
        'step1': [('작업명', a.job_title or job.name),
                  ('작업내용(단위작업명)', '  '.join(f'{i}) {t.get("name")}' for i, t in enumerate(tasks, 1)) or '')],
        'tasks': t_rows, 'worker': getattr(a, 'worker_name', None) or '', 'hazards': h_rows, 'ergo': e_rows,
        'summary': summary, 'symptom': symptom,
        'ustats': [[k, v['n'], v['a'], v['b'], v['score'], ', '.join(f'{i}호({c})' for i, c in v.get('sug', [])) or '-'] for k, v in (js.get('ustats') or {}).items()],
    }


# ══════════════════════════════════════════════════════════════════════
#  Excel
# ══════════════════════════════════════════════════════════════════════
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference

_thin = Side(style='thin', color='A0A7B4')
BORDER = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)
HFILL = PatternFill('solid', fgColor='1D2330')
SFILL = PatternFill('solid', fgColor='EEF1F4')
RFILL = PatternFill('solid', fgColor='FDE8E8')
YFILL = PatternFill('solid', fgColor='FEF0C7')
WRAP = Alignment(wrap_text=True, vertical='center')
CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)
F_TITLE = Font(name='맑은 고딕', size=16, bold=True)
F_H = Font(name='맑은 고딕', size=10, bold=True, color='FFFFFF')
F_B = Font(name='맑은 고딕', size=10)
F_BB = Font(name='맑은 고딕', size=10, bold=True)
F_SEC = Font(name='맑은 고딕', size=12, bold=True, color='C8102E')


class XW:
    """시트에 위에서 아래로 블록을 쌓는 작은 헬퍼"""
    def __init__(self, ws, widths):
        self.ws, self.r = ws, 1
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        self.ncol = len(widths)

    def title(self, text):
        c = self.ws.cell(self.r, 1, text)
        c.font = F_TITLE
        self.r += 2

    def sec(self, text):
        c = self.ws.cell(self.r, 1, text)
        c.font = F_SEC
        self.r += 1

    def para(self, text, bold=False):
        self.ws.merge_cells(start_row=self.r, start_column=1, end_row=self.r, end_column=self.ncol)
        c = self.ws.cell(self.r, 1, text)
        c.font = F_BB if bold else F_B
        c.alignment = Alignment(wrap_text=True, vertical='top')
        self.ws.row_dimensions[self.r].height = max(18, 15 * (len(text) // 95 + 1))
        self.r += 1

    def header(self, vals, start=1):
        for i, v in enumerate(vals):
            c = self.ws.cell(self.r, start + i, v)
            c.font, c.fill, c.alignment, c.border = F_H, HFILL, CENTER, BORDER
        self.r += 1

    def row(self, vals, start=1, fill=None, bold=False, center_from=None):
        for i, v in enumerate(vals):
            c = self.ws.cell(self.r, start + i, v)
            c.font = F_BB if bold else F_B
            c.border = BORDER
            c.alignment = CENTER if (center_from is not None and i >= center_from) else WRAP
            if fill:
                c.fill = fill
        self.r += 1

    def kv(self, k, v, span=None):
        span = span or self.ncol
        a = self.ws.cell(self.r, 1, k)
        a.font, a.fill, a.border, a.alignment = F_BB, SFILL, BORDER, WRAP
        self.ws.merge_cells(start_row=self.r, start_column=2, end_row=self.r, end_column=span)
        b = self.ws.cell(self.r, 2, v)
        b.font, b.alignment = F_B, Alignment(wrap_text=True, vertical='center')
        for col in range(2, span + 1):
            self.ws.cell(self.r, col).border = BORDER
        self.ws.row_dimensions[self.r].height = max(20, 15 * (len(str(v or '')) // 80 + 1))
        self.r += 1

    def gap(self, n=1):
        self.r += n


def build_excel(A, company, site):
    r = A['round']
    wb = Workbook()

    # ── 표지 ──
    ws = wb.active
    ws.title = '표지'
    for col, w in zip('ABCDEFGH', [4, 14, 14, 14, 14, 14, 14, 4]):
        ws.column_dimensions[col].width = w
    ws.merge_cells('B4:G6')
    c = ws['B4']
    c.value = f"{r.title if r else ''}\n결과보고서"
    c.font = Font(name='맑은 고딕', size=22, bold=True)
    c.alignment = CENTER
    info = [('사업장', f'{company} {site}'), ('조사구분', r.kind if r else ''), ('조사기간', period(r)),
            ('조사자', f"{r.investigator or ''} {('(' + r.agency + ')') if r and r.agency else ''}" if r else ''),
            ('작성일', (datetime.utcnow() + KST).strftime('%Y.%m.%d'))]
    for i, (k, v) in enumerate(info):
        rr = 10 + i
        ws.merge_cells(start_row=rr, start_column=2, end_row=rr, end_column=3)
        ws.merge_cells(start_row=rr, start_column=4, end_row=rr, end_column=7)
        ws.cell(rr, 2, k).font = F_BB
        ws.cell(rr, 2).fill = SFILL
        ws.cell(rr, 4, v).font = F_B
        for col in range(2, 8):
            ws.cell(rr, col).border = BORDER
            ws.cell(rr, col).alignment = CENTER
        ws.row_dimensions[rr].height = 24
    # 결재란
    for i, t in enumerate(['담당', '검토', '승인']):
        c1 = ws.cell(18, 5 + i, t)
        c1.font, c1.fill, c1.alignment, c1.border = F_BB, SFILL, CENTER, BORDER
        ws.cell(19, 5 + i).border = BORDER
    ws.row_dimensions[19].height = 50
    ws.merge_cells('B24:G24')
    ws['B24'] = f'{company}'
    ws['B24'].font = Font(name='맑은 고딕', size=16, bold=True)
    ws['B24'].alignment = CENTER

    # ── 요약 (1쪽, 경영진 보고용) ──
    SD = summary_data(A)
    x = XW(wb.create_sheet('요약'), [18, 22, 30, 60])
    x.title('유해요인조사 결과 요약')
    for k, v in SD['overview']:
        x.kv(k, v)
    x.gap()
    x.sec('종합판정 (유해요인조사표 × 증상조사표)')
    x.header(['판정', '작업', '의미', ''])
    for a_, b_, c_ in SD['quad']:
        x.row([a_, b_, c_, ''], fill=RFILL if a_.startswith('①') else (YFILL if a_.startswith('②') else None))
    x.gap()
    x.sec('개선 우선순위 TOP 5')
    x.header(['순위', '판정', '작업', '근거'])
    for row in SD['top']:
        x.row(row, center_from=None)
    if not SD['top']:
        x.para('해당 없음')

    # ── 1. 조사개요 ──
    x = XW(wb.create_sheet('1.조사개요'), [22, 20, 20, 20, 20, 20])
    x.title('1. 조사 개요')
    x.kv('사업장', f'{company} {site}')
    x.kv('조사구분', r.kind if r else '')
    x.kv('조사기간', period(r))
    x.kv('조사자(기관)', f"{r.investigator or ''} {r.agency or ''}".strip() if r else '')
    x.kv('법적 근거', LEGAL)
    x.kv('조사 대상', f"작업 {len([j for j in A['job_summary'] if j['job'].active])}개, 대상 근로자 {A['target_workers'] or '-'}명, "
                     f"증상조사 응답 {A['n']}명" + (f" (응답률 {round(A['n'] * 100 / A['target_workers'], 1)}%)" if A['target_workers'] else ''))
    for i, (k, v) in enumerate(METHOD, 1):
        x.kv(f'조사방법 {i}', f'{k}: {v}')
    x.kv('판정기준', '\n'.join(S.JUDGE_TEXT.values()))
    if r and r.overview:
        x.kv('기타', r.overview)

    # ── 2. 체크리스트 ──
    x = XW(wb.create_sheet('2.부담작업해당여부'), [14, 12, 18, 24, 8] + [5] * 11 + [10])
    x.title('2. 근골격계부담작업 체크리스트 (단위작업별)')
    x.para('편람 기준: 단위작업명과 근로자 수를 적고, 부담작업에 해당하면 ○, 아니면 × — 11가지 중 하나라도 ○이면 유해요인조사 대상')
    x.header(['부서', '공정', '작업명', '단위작업명', '근로자수'] + [f'{n}호' for n, _, _ in S.BURDEN_WORKS] + ['해당여부'])
    hit = {n: 0 for n, _, _ in S.BURDEN_WORKS}
    for js in A['job_summary']:
        a = js['assess']
        units = jl(a.checklist, []) if a else []
        if not units:
            x.row([js['job'].dept, js['job'].process or '', js['job'].name, '(미조사)' if not a else '-', js['job'].workers or ''] + [''] * 11 + ['' if not a else '비해당'], center_from=4)
            continue
        for u in units:
            items = [int(i) for i in u.get('items', [])]
            for i in items:
                hit[i] += 1
            x.row([js['job'].dept, js['job'].process or '', js['job'].name, u.get('unit'), unit_workers(a).get(u.get('unit')) or ''] +
                  ['○' if n in items else '×' for n, _, _ in S.BURDEN_WORKS] + ['해당' if items else '비해당'],
                  center_from=4, fill=YFILL if items else None)
    x.row(['', '', '', '호별 해당 단위작업 수', ''] + [hit[n] or '' for n, _, _ in S.BURDEN_WORKS] + [''], bold=True, center_from=4, fill=SFILL)
    x.gap()
    x.sec('※ 근골격계부담작업 (고용노동부 고시)')
    for n, t, p in S.BURDEN_WORKS:
        x.para(f'제{n}호 [{p}] {t}')

    # ── 3. 유해요인조사표 [별지 제1호서식] ──
    x = XW(wb.create_sheet('3.유해요인조사표'), [24, 13, 22, 26, 30, 16])
    x.title('3. 유해요인조사표 (작업별, [별지 제1호서식])')
    if not A['assessments']:
        x.para('작성된 유해요인조사표가 없습니다.')
    for js in A['job_summary']:
        a = js['assess']
        if not a:
            continue
        F = f1_sections(a, js, r)
        x.sec(f"■ {js['job'].dept} / {js['job'].name}")
        x.para('가. 조사 개요', bold=True)
        for k, v in F['overview']:
            x.kv(k, v)
        x.para('나. 작업장 상황 조사', bold=True)
        for k, v in F['site']:
            x.kv(k, v)
        x.para('다. 작업조건 조사 — 1단계: 작업별 주요 작업내용', bold=True)
        for k, v in F['step1']:
            x.kv(k, v)
        x.para('2단계: 작업별 작업부하 및 작업빈도', bold=True)
        x.header(['단위작업명', '부담작업(호)', '작업부하(A)', '작업빈도(B)', '총점수(A×B)', ''])
        for row in F['tasks']:
            x.row(row + [''], center_from=1)
        if F['ustats']:
            x.para('2단계 참고: 근로자 모바일 응답(면담) 집계', bold=True)
            x.header(['단위작업명', '응답(명)', '평균 작업부하(A)', '평균 작업빈도(B)', '평균 총점수', '부담작업 후보(응답자 수)'])
            for row in F['ustats']:
                x.row(row, center_from=1)
        x.para('3단계: 유해요인평가' + (f"  (근로자명: {F['worker']})" if F['worker'] else ''), bold=True)
        x.header(['단위작업명', '부담작업(호)', '유해요인', '발생원인', '', '비고'])
        for h in F['hazards']:
            x.row(h[:4] + [''] + h[4:])
            x.ws.merge_cells(start_row=x.r - 1, start_column=4, end_row=x.r - 1, end_column=5)
        if not F['hazards']:
            x.row(['-', '', '', '', '', ''])
        if F['ergo']:
            x.para('3-1. 인간공학적 작업·분석평가 결과', bold=True)
            x.header(['단위작업명', '평가도구', '평가결과', '판정', '', ''])
            for e in F['ergo']:
                x.row(e + ['', ''])
        x.para('라. 유해요인 조사 결과', bold=True)
        x.kv('결과요약', F['summary'])
        x.kv('증상조사(별지 제2호)', F['symptom'])
        x.gap()

    # ── 4. 증상조사 결과 ──
    ws4 = wb.create_sheet('4.증상조사결과')
    x = XW(ws4, [22, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12])
    x.title('4. 근골격계질환 증상조사 결과')
    x.para(f"응답자 {A['n']}명 · " + ' / '.join(S.JUDGE_TEXT.values()))
    x.gap()
    x.sec('4-1. 응답자 일반적 특성')
    for title, rows in list(A['general'].items()) + list(A['general_extra'].items()):
        x.header([title, '인원(명)', '비율(%)'])
        for k, c, p in rows:
            x.row([k, c, p], center_from=1)
        x.gap()
    x.sec('4-2. 신체부위별 증상 현황')
    x.header(['부위', '증상호소(명)', '(%)', '관리대상(명)', '(%)', '통증호소(명)', '(%)'])
    chart_top = x.r
    for p in A['parts']:
        x.row([p['label'], p['complaint'], p['complaint_p'], p['manage'], p['manage_p'], p['pain'], p['pain_p']], center_from=1)
    x.row(['전체(인원 기준)', A['persons']['complaint'], A['persons_p']['complaint'], A['persons']['manage'], A['persons_p']['manage'],
           A['persons']['pain'], A['persons_p']['pain']], bold=True, fill=SFILL, center_from=1)
    ch = BarChart()
    ch.type = 'col'
    ch.title = '신체부위별 증상호소율(%)'
    ch.y_axis.title = '%'
    for col in (3, 5, 7):
        ch.add_data(Reference(ws4, min_col=col, min_row=chart_top - 1, max_row=chart_top + 5), titles_from_data=True)
    ch.set_categories(Reference(ws4, min_col=1, min_row=chart_top, max_row=chart_top + 5))
    for s, nm in zip(ch.series, ['증상호소', '관리대상', '통증호소']):
        from openpyxl.chart.series import SeriesLabel
        s.tx = SeriesLabel(v=nm)
    ch.height, ch.width = 7.5, 16
    ws4.add_chart(ch, f'I{chart_top - 1}')
    x.gap()
    x.sec('4-3. 부위별 통증 특성 (응답 인원)')
    x.header(['부위'] + [S.label('dur', d) for d in S.DURATION])
    for p in A['parts']:
        x.row([p['label']] + [p['dur'][d] for d in S.DURATION], center_from=1)
    x.header(['부위'] + [S.label('int', d) for d in S.INTENSITY])
    for p in A['parts']:
        x.row([p['label']] + [p['int'][d] for d in S.INTENSITY], center_from=1)
    x.header(['부위'] + [S.label('freq', d) for d in S.FREQUENCY])
    for p in A['parts']:
        x.row([p['label']] + [p['freq'][d] for d in S.FREQUENCY], center_from=1)
    x.gap()
    for sec_title, grp in [('4-4. 부서별 증상 현황', A['by_dept']), ('4-5. 작업별 증상 현황', A['by_job'])]:
        x.sec(sec_title)
        x.header(['구분', '응답(명)', '증상호소', '(%)', '관리대상', '(%)', '통증호소', '(%)'] + [])
        for g in grp:
            x.row([g['key'], g['n'], g['complaint'], g['complaint_p'], g['manage'], g['manage_p'], g['pain'], g['pain_p']], center_from=1)
        x.gap()
    x.sec('4-6. 통증호소자 명단 (대외비) — 전체 결과표는 4-1 시트')
    x.header(['성명', '부서', '작업', '연령', '해당 부위', '사후관리', '', '', '', '', ''])
    for rr in A['pain_list']:
        parts = ', '.join(S.label('part', p) for p, v in rr.part_levels.items() if v == 'pain')
        x.row([rr.name, rr.dept, rr.job_label, rr.age, parts, rr.followup or '미조치'], fill=RFILL)
    if not A['pain_list']:
        x.para('해당자 없음')
    x.gap()
    x.sec('4-7. 관리대상자 명단 (통증호소자 제외)')
    x.header(['성명', '부서', '작업', '연령', '해당 부위', '사후관리'])
    for rr in A['manage_list']:
        parts = ', '.join(S.label('part', p) for p, v in rr.part_levels.items() if v == 'manage')
        x.row([rr.name, rr.dept, rr.job_label, rr.age, parts, rr.followup or ''], fill=YFILL)
    if not A['manage_list']:
        x.para('해당자 없음')

    # ── 4-1. KOSHA 결과표 ──
    wsk = wb.create_sheet('4-1.증상조사결과표')
    x = XW(wsk, [5, 10, 7, 8, 11, 11, 16, 22, 11, 10, 10, 10, 10, 10, 10, 11])
    x.title('근골격계질환 증상조사표 결과')
    x.para(KOSHA_NOTE)
    x.header(KOSHA_HEAD)
    for row in kosha_rows(A):
        x.row(row, center_from=2)
        for ci in range(10, 17):
            c = wsk.cell(x.r - 1, ci)
            if c.value == '통증호소자':
                c.fill, c.font = RFILL, F_BB
            elif c.value == '관리대상자':
                c.fill = YFILL
    wsk.freeze_panes = 'C' + str(x.r - len(A['responses']))
    wsk.page_setup.orientation = 'landscape'

    # ── 5. 종합판정 ──
    basis, fit, pers = judge_rows(A)
    x = XW(wb.create_sheet('5.종합판정'), [14, 24, 10, 9, 8, 13, 9, 60])
    x.title('5. 유해요인조사표 × 증상조사표 종합판정')
    x.sec('5-1. 판정 기준')
    for k, v in JUDGE_BASIS:
        x.kv(k, v)
    x.gap()
    x.sec('5-2. 작업별 판정 및 근거')
    x.header(['판정', '작업', '부담작업(호)', 'A×B 최고', '응답', '관리대상 이상', '통증호소', '근거'])
    for row in basis:
        x.row(row, center_from=2, fill=RFILL if row[0].startswith('①') else (YFILL if row[0].startswith('②') else None))
    x.gap()
    x.sec('5-3. 신체부위 부합성 (● 부담작업 관련 부위 · 숫자 관리대상/통증호소 · ◎ 부합 · △ 조사표 외 부위)')
    x.header(['작업', '부담작업'] + [S.label('part', q) for q in S.PARTS])
    for row in fit:
        x.row(row[:-1], center_from=1)
    x.gap()
    x.sec('5-4. 개인요인 확인 (관리대상자·통증호소자)')
    x.header(['성명', '판정', '작업', '증상 부위', '작업 관련 부위', '개인요인', '판단', '사후관리'])
    for row in pers:
        x.row(row)
    if not pers:
        x.para('해당자 없음')
    x.ws.page_setup.orientation = 'landscape'

    # ── 6. 개선계획서 ──
    x = XW(wb.create_sheet('6.개선계획서'), [8, 14, 16, 34, 26, 34, 14, 12, 10, 10, 14])
    x.title('6. 작업환경 개선계획서')
    x.header(['우선순위', '공정명', '작업명', '문제점(유해요인의 원인)', '근로자 의견', '개선방안', '추진일정', '개선비용(천원)', '총점수', '증상호소', '진행상태'])
    smap = {j['job'].id: j for j in A['job_summary']}
    for m in A['improvements']:
        js = smap.get(m.job_id, {})
        x.row([m.priority, (m.job.process or m.job.dept) if m.job else '', m.job.name if m.job else '', m.problem, m.opinion, m.plan,
               m.schedule, m.cost, js.get('score') or '', js.get('symptom', ''),
               (m.status or '') + (f' ({fmt_date(m.done_date)})' if m.done_date else '')])
    if not A['improvements']:
        x.para('등록된 개선계획이 없습니다.')
    x.gap()
    x.para('위 작업환경 개선계획을 확인합니다.   사업주:                 (인)      근로자 대표:                 (인)', bold=True)
    x.gap()
    x.sec('개선 우선순위 선정 결과(자동)')
    x.header(['순위', '부서', '작업명', '부담작업(호)', 'A×B 최고', '관리대상', '통증호소', '', '', '', ''])
    for j in A['ranked']:
        x.row([j['rank'], j['job'].dept, j['job'].name, ', '.join(map(str, j['items'])) or '비해당', j['score'], j['manage'], j['pain']], center_from=3)

    # ── 7. 종합의견 ──
    x = XW(wb.create_sheet('7.종합의견'), [110])
    x.title('7. 종합 의견')
    for i, p in enumerate(conclusion_text(A), 1):
        x.para(f'{i}. {p}')

    # ── 원자료 ──
    _raw_sheet(wb.create_sheet('원자료'), A)
    _unit_sheet(wb.create_sheet('작업별응답(원자료)'), A)

    for w in wb.worksheets:
        w.sheet_view.showGridLines = w.title == '원자료'
        w.page_setup.paperSize = 9
        w.page_setup.fitToWidth = 1
        w.page_setup.fitToHeight = 0
        w.sheet_properties.pageSetUpPr.fitToPage = True
    wb['2.부담작업해당여부'].page_setup.orientation = 'landscape'
    wb['6.개선계획서'].page_setup.orientation = 'landscape'
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _raw_sheet(ws, A):
    head, rows = raw_rows(A)
    for i, h in enumerate(head, 1):
        c = ws.cell(1, i, h)
        c.font, c.fill, c.alignment, c.border = F_H, HFILL, CENTER, BORDER
        ws.column_dimensions[get_column_letter(i)].width = max(8, min(30, len(h) * 1.8))
    for ri, row in enumerate(rows, 2):
        for ci, v in enumerate(row, 1):
            c = ws.cell(ri, ci, v)
            c.font, c.border = F_B, BORDER
            if isinstance(v, str) and v in ('통증호소',):
                c.fill = RFILL
    ws.freeze_panes = 'E2'
    ws.auto_filter.ref = f'A1:{get_column_letter(len(head))}{max(1, len(rows) + 1)}'


UNIT_HEAD = ['응답번호', '성명', '부서', '작업', '단위작업', '작업부하(A)', '작업빈도(B)', '총점수(A×B)', '증상 전체결과',
             '하루 시간', '무게', '하루 횟수', '자세·동작', '가장 힘든 점', '구체 내용', '부담작업 후보(호)']


def _unit_sheet(ws, A):
    """응답자 × 선택 작업 1행 — 피벗·통계용"""
    for i, h in enumerate(UNIT_HEAD, 1):
        c = ws.cell(1, i, h)
        c.font, c.fill, c.alignment, c.border = F_H, HFILL, CENTER, BORDER
        ws.column_dimensions[get_column_letter(i)].width = [9, 10, 14, 18, 26, 11, 11, 12, 12, 11, 12, 14, 34, 16, 34, 14][i - 1]
    for ri, row in enumerate(A.get('unit_rows', []), 2):
        for ci, v in enumerate(row, 1):
            c = ws.cell(ri, ci, v)
            c.font, c.border = F_B, BORDER
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = f'A1:{get_column_letter(len(UNIT_HEAD))}{max(1, len(A.get("unit_rows", [])) + 1)}'


def build_raw_excel(A):
    wb = Workbook()
    _raw_sheet(wb.active, A)
    wb.active.title = '증상조사 원자료'
    _unit_sheet(wb.create_sheet('작업별 부하·빈도 응답'), A)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ══════════════════════════════════════════════════════════════════════
#  PDF (reportlab)
# ══════════════════════════════════════════════════════════════════════
def _font():
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    base = os.path.dirname(os.path.abspath(__file__))
    fdir = os.path.join(base, 'static', 'fonts')
    cands = [(os.path.join(fdir, 'NanumGothic.ttf'), os.path.join(fdir, 'NanumGothicBold.ttf')),
             ('C:/Windows/Fonts/malgun.ttf', 'C:/Windows/Fonts/malgunbd.ttf'),
             (os.path.join(fdir, 'NanumGothicBold.ttf'), os.path.join(fdir, 'NanumGothicBold.ttf')),   # 서버: 동봉 폰트
             ('C:/Windows/Fonts/malgun.ttf', 'C:/Windows/Fonts/malgunbd.ttf'),
             ('/usr/share/fonts/truetype/nanum/NanumGothic.ttf', '/usr/share/fonts/truetype/nanum/NanumGothicBold.ttf')]
    for reg, bold in cands:
        if os.path.exists(reg):
            try:
                pdfmetrics.registerFont(TTFont('KR', reg))
                pdfmetrics.registerFont(TTFont('KRB', bold if os.path.exists(bold) else reg))
                return 'KR', 'KRB'
            except Exception:
                pass
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont   # 폰트 파일이 없을 때 (서버 등)
    pdfmetrics.registerFont(UnicodeCIDFont('HYGothic-Medium'))
    return 'HYGothic-Medium', 'HYGothic-Medium'


def build_pdf(A, company, site, photos=None):
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.lib.units import mm
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak,
                                    Image, KeepTogether)
    from reportlab.graphics.shapes import Drawing
    from reportlab.graphics.charts.barcharts import VerticalBarChart
    from reportlab.graphics.charts.legends import Legend
    from xml.sax.saxutils import escape

    FN, FB = _font()
    r = A['round']
    photos = photos or {}
    st = {
        'title': ParagraphStyle('t', fontName=FB, fontSize=22, leading=30, alignment=1),
        'h1': ParagraphStyle('h1', fontName=FB, fontSize=14, leading=20, spaceBefore=6, spaceAfter=6, textColor=colors.HexColor('#C8102E')),
        'h2': ParagraphStyle('h2', fontName=FB, fontSize=11, leading=16, spaceBefore=8, spaceAfter=4),
        'b': ParagraphStyle('b', fontName=FN, fontSize=9, leading=13),
        'c': ParagraphStyle('c', fontName=FN, fontSize=8.5, leading=11),
        'cc': ParagraphStyle('cc', fontName=FN, fontSize=8.5, leading=11, alignment=1),
        'ch': ParagraphStyle('ch', fontName=FB, fontSize=8.5, leading=11, alignment=1, textColor=colors.white),
    }
    P = lambda t, s='c': Paragraph(escape(str(t if t is not None else '')).replace('\n', '<br/>'), st[s])
    W = A4[0] - 30 * mm

    def tbl(head, rows, widths=None, center_from=1, row_fills=None, tw=None):
        data = [[P(h, 'ch') for h in head]] + [[P(v, 'cc' if i >= center_from else 'c') for i, v in enumerate(rw)] for rw in rows]
        if widths:
            tot = sum(widths)
            widths = [(tw or W) * w / tot for w in widths]
        elif tw:
            widths = [tw / len(head)] * len(head)
        t = Table(data, colWidths=widths, repeatRows=1)
        sty = [('GRID', (0, 0), (-1, -1), 0.4, colors.HexColor('#A0A7B4')),
               ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1D2330')),
               ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
               ('TOPPADDING', (0, 0), (-1, -1), 2), ('BOTTOMPADDING', (0, 0), (-1, -1), 2)]
        for i, f in (row_fills or {}).items():
            sty.append(('BACKGROUND', (0, i + 1), (-1, i + 1), colors.HexColor(f)))
        t.setStyle(TableStyle(sty))
        return t

    def kv(rows):
        data = [[P(k, 'c'), P(v, 'c')] for k, v in rows]
        t = Table(data, colWidths=[W * 0.22, W * 0.78])
        t.setStyle(TableStyle([('GRID', (0, 0), (-1, -1), 0.4, colors.HexColor('#A0A7B4')),
                               ('BACKGROUND', (0, 0), (0, -1), colors.HexColor('#EEF1F4')),
                               ('VALIGN', (0, 0), (-1, -1), 'MIDDLE')]))
        return t

    def footer(canvas, doc):
        canvas.saveState()
        canvas.setFont(FN, 8)
        canvas.setFillColor(colors.HexColor('#667085'))
        canvas.drawString(15 * mm, 10 * mm, f"{company} · {r.title if r else ''}")
        canvas.drawRightString(A4[0] - 15 * mm, 10 * mm, f'- {doc.page} -')
        canvas.restoreState()

    el = []

    def newpage():
        while el and isinstance(el[-1], Spacer):
            el.pop()
        el.append(PageBreak())
    # 표지
    el += [Spacer(1, 60 * mm), P(f"{r.title if r else ''}", 'title'), Spacer(1, 4 * mm), P('결과보고서', 'title'), Spacer(1, 30 * mm)]
    cover = [('사업장', f'{company} {site}'), ('조사구분', r.kind if r else ''), ('조사기간', period(r)),
             ('조사자(기관)', f"{r.investigator or ''} {r.agency or ''}".strip() if r else ''),
             ('작성일', (datetime.utcnow() + KST).strftime('%Y.%m.%d'))]
    el.append(kv(cover))
    el.append(Spacer(1, 15 * mm))
    sign = Table([[P('담당', 'cc'), P('검토', 'cc'), P('승인', 'cc')], ['', '', '']], colWidths=[25 * mm] * 3, rowHeights=[7 * mm, 18 * mm], hAlign='RIGHT')
    sign.setStyle(TableStyle([('GRID', (0, 0), (-1, -1), 0.5, colors.black), ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#EEF1F4'))]))
    el += [sign, PageBreak()]

    # 요약 (1쪽)
    SD = summary_data(A)
    el += [P('결과 요약', 'h1'), kv(SD['overview']), Spacer(1, 4 * mm), P('종합판정 (유해요인조사표 × 증상조사표)', 'h2')]
    qf = {i: ('#FDE8E8' if row[0].startswith('①') else '#FEF0C7' if row[0].startswith('②') else '#FFFFFF') for i, row in enumerate(SD['quad'])}
    el.append(tbl(['판정', '작업', '의미'], [list(q) for q in SD['quad']] or [['-', '-', '-']], [16, 34, 50], center_from=9, row_fills=qf))
    el += [Spacer(1, 4 * mm), P('개선 우선순위 TOP 5', 'h2'),
           tbl(['순위', '판정', '작업', '근거'], SD['top'] or [['', '', '해당 없음', '']], [6, 14, 24, 56], center_from=9)]
    newpage()

    # 1. 개요
    el.append(P('1. 조사 개요', 'h1'))
    ov = [('사업장', f'{company} {site}'), ('조사구분', r.kind if r else ''), ('조사기간', period(r)),
          ('조사자(기관)', f"{r.investigator or ''} {r.agency or ''}".strip() if r else ''), ('법적 근거', LEGAL),
          ('조사 대상', f"작업 {len([j for j in A['job_summary'] if j['job'].active])}개, 대상 근로자 {A['target_workers'] or '-'}명, 증상조사 응답 {A['n']}명")]
    ov += [(f'조사방법 {i}', f'{k}: {v}') for i, (k, v) in enumerate(METHOD, 1)]
    ov.append(('판정기준', '\n'.join(S.JUDGE_TEXT.values())))
    if r and r.overview:
        ov.append(('기타', r.overview))
    el.append(kv(ov))

    # 2. 체크리스트
    newpage()
    el += [P('2. 근골격계부담작업 체크리스트 (단위작업별)', 'h1')]
    rows, fills = [], {}
    for js in A['job_summary']:
        a = js['assess']
        units = jl(a.checklist, []) if a else []
        if not units:
            rows.append([js['job'].dept, js['job'].name, '(미조사)' if not a else '-', js['job'].workers or ''] + [''] * 11 + ['' if not a else '비해당'])
            continue
        for u in units:
            items = [int(i) for i in u.get('items', [])]
            if items:
                fills[len(rows)] = '#FEF0C7'
            rows.append([js['job'].dept, js['job'].name, u.get('unit'), unit_workers(a).get(u.get('unit')) or ''] + ['○' if n in items else '×' for n, _, _ in S.BURDEN_WORKS] + ['해당' if items else '비해당'])
    el.append(P('단위작업별로 부담작업에 해당하면 ○, 아니면 × (11가지 중 하나라도 ○이면 유해요인조사 대상)', 'c'))
    el.append(tbl(['부서', '작업명', '단위작업', '인원'] + [str(n) for n, _, _ in S.BURDEN_WORKS] + ['해당'], rows,
                  [9, 11, 13, 4.5] + [4.1] * 11 + [5.5], center_from=3, row_fills=fills))
    el.append(Spacer(1, 3 * mm))
    for n, t, p in S.BURDEN_WORKS:
        el.append(P(f'제{n}호 [{p}] {t}', 'c'))

    # 3. 유해요인조사표 [별지 제1호서식]
    newpage()
    el += [P('3. 유해요인조사표 (작업별, [별지 제1호서식])', 'h1')]
    if not A['assessments']:
        el.append(P('작성된 유해요인조사표가 없습니다.', 'b'))
    for js in A['job_summary']:
        a = js['assess']
        if not a:
            continue
        F = f1_sections(a, js, r)
        el.append(KeepTogether([P(f"■ {js['job'].dept} / {js['job'].name}", 'h2'), P('가. 조사 개요', 'b'), kv(F['overview']),
                                Spacer(1, 2 * mm), P('나. 작업장 상황 조사', 'b'), kv(F['site'])]))
        el.append(KeepTogether([Spacer(1, 2 * mm), P('다. 작업조건 조사 — 1단계: 작업별 주요 작업내용', 'b'), kv(F['step1']),
                                Spacer(1, 2 * mm), P('2단계: 작업별 작업부하 및 작업빈도', 'b'),
                                tbl(['단위작업명', '부담작업(호)', '작업부하(A)', '작업빈도(B)', '총점수(A×B)'],
                                    F['tasks'] or [['-', '', '', '', '']], [28, 15, 17, 29, 11])]))
        if F['ustats']:
            el.append(KeepTogether([Spacer(1, 2 * mm), P('2단계 참고: 근로자 모바일 응답(면담) 집계', 'b'),
                                    tbl(['단위작업명', '응답', '평균 A', '평균 B', '평균 총점', '부담작업 후보(응답자 수)'], F['ustats'], [28, 8, 10, 10, 10, 34])]))
        el.append(KeepTogether([Spacer(1, 2 * mm), P('3단계: 유해요인평가' + (f"  (근로자명: {F['worker']})" if F['worker'] else ''), 'b'),
                                tbl(['단위작업명', '부담작업(호)', '유해요인', '발생원인', '비고'],
                                    F['hazards'] or [['-', '', '', '', '']], [19, 13, 18, 36, 14], center_from=1)]))
        if F['ergo']:
            el += [Spacer(1, 2 * mm), P('3-1. 인간공학적 작업·분석평가 결과', 'b'),
                   tbl(['단위작업명', '평가도구', '평가결과', '판정'], F['ergo'], [25, 20, 30, 25])]
        el.append(KeepTogether([Spacer(1, 2 * mm), P('라. 유해요인 조사 결과', 'b'),
                                kv([('결과요약', F['summary']), ('증상조사(별지 제2호)', F['symptom'])])]))
        imgs = []
        for cap, data in photos.get(a.id, [])[:4]:
            try:
                from PIL import Image as PI
                im = PI.open(io.BytesIO(data))
                w, h = im.size
                iw = W / 2 - 4 * mm
                ih = min(iw * h / w, 70 * mm)
                iw = ih * w / h
                imgs.append([Image(io.BytesIO(data), width=iw, height=ih), P(cap, 'cc')])
            except Exception:
                pass
        if imgs:
            cells = [imgs[i:i + 2] for i in range(0, len(imgs), 2)]
            rows_ = []
            for pair in cells:
                rows_.append([c[0] for c in pair] + [''] * (2 - len(pair)))
                rows_.append([c[1] for c in pair] + [''] * (2 - len(pair)))
            el.append(Table(rows_, colWidths=[W / 2, W / 2]))
        el.append(Spacer(1, 5 * mm))

    # 4. 증상조사 결과
    newpage()
    el += [P('4. 근골격계질환 증상조사 결과', 'h1'),
           P(f"응답자 {A['n']}명. " + ' / '.join(S.JUDGE_TEXT.values()), 'b'), P('4-1. 응답자 일반적 특성', 'h2')]
    gen = list(A['general'].items())
    pairs = [gen[i:i + 2] for i in range(0, len(gen), 2)]
    for pr in pairs:
        cells = []
        for title, rows in pr:
            cells.append(tbl([title, '명', '%'], [[k, c, p] for k, c, p in rows], [50, 20, 20], tw=W / 2 - 6))
        if len(cells) == 1:
            cells.append('')
        t = Table([cells], colWidths=[W / 2, W / 2])
        t.setStyle(TableStyle([('VALIGN', (0, 0), (-1, -1), 'TOP'), ('LEFTPADDING', (0, 0), (-1, -1), 2), ('RIGHTPADDING', (0, 0), (-1, -1), 2)]))
        el += [t, Spacer(1, 2 * mm)]
    el.append(P('4-2. 신체부위별 증상 현황', 'h2'))
    prow = [[p['label'], p['complaint'], p['complaint_p'], p['manage'], p['manage_p'], p['pain'], p['pain_p']] for p in A['parts']]
    prow.append(['전체(인원)', A['persons']['complaint'], A['persons_p']['complaint'], A['persons']['manage'], A['persons_p']['manage'],
                 A['persons']['pain'], A['persons_p']['pain']])
    el.append(tbl(['부위', '증상호소(명)', '%', '관리대상(명)', '%', '통증호소(명)', '%'], prow, [22, 14, 10, 14, 10, 14, 10], row_fills={len(prow) - 1: '#EEF1F4'}))
    # 막대그래프
    d = Drawing(W, 62 * mm)
    bc = VerticalBarChart()
    bc.x, bc.y, bc.width, bc.height = 12 * mm, 10 * mm, W - 50 * mm, 45 * mm
    bc.data = [[p['complaint_p'] for p in A['parts']], [p['manage_p'] for p in A['parts']], [p['pain_p'] for p in A['parts']]]
    bc.categoryAxis.categoryNames = [p['label'] for p in A['parts']]
    bc.categoryAxis.labels.fontName = FN
    bc.categoryAxis.labels.fontSize = 7.5
    bc.valueAxis.labels.fontName = FN
    bc.valueAxis.valueMin = 0
    bc.valueAxis.valueMax = max(10, max([max(s) for s in bc.data] or [0]) * 1.15)
    for i, c in enumerate(['#A4BCFD', '#FDB022', '#C8102E']):
        bc.bars[i].fillColor = colors.HexColor(c)
        bc.bars[i].strokeColor = None
    lg = Legend()
    lg.x, lg.y = W - 34 * mm, 45 * mm
    lg.fontName, lg.fontSize = FN, 8
    lg.colorNamePairs = [(colors.HexColor('#A4BCFD'), '증상호소%'), (colors.HexColor('#FDB022'), '관리대상%'), (colors.HexColor('#C8102E'), '통증호소%')]
    d.add(bc)
    d.add(lg)
    el.append(d)
    el.append(P('4-3. 부위별 통증 특성 (응답 인원)', 'h2'))
    el.append(tbl(['부위'] + [S.label('dur', x) for x in S.DURATION], [[p['label']] + [p['dur'][x] for x in S.DURATION] for p in A['parts']]))
    el.append(Spacer(1, 2 * mm))
    el.append(tbl(['부위'] + [S.label('int', x) for x in S.INTENSITY], [[p['label']] + [p['int'][x] for x in S.INTENSITY] for p in A['parts']]))
    el.append(Spacer(1, 2 * mm))
    el.append(tbl(['부위'] + [S.label('freq', x) for x in S.FREQUENCY], [[p['label']] + [p['freq'][x] for x in S.FREQUENCY] for p in A['parts']]))
    for title, grp in [('4-4. 부서별 증상 현황', A['by_dept']), ('4-5. 작업별 증상 현황', A['by_job'])]:
        el.append(P(title, 'h2'))
        el.append(tbl(['구분', '응답', '증상호소', '%', '관리대상', '%', '통증호소', '%'],
                      [[g['key'], g['n'], g['complaint'], g['complaint_p'], g['manage'], g['manage_p'], g['pain'], g['pain_p']] for g in grp] or [['응답 없음'] + [''] * 7],
                      [34, 9, 11, 9, 11, 9, 11, 9]))
    el.append(P('4-6. 증상조사 결과표 (KOSHA H-30-2008 양식)', 'h2'))
    el.append(P(KOSHA_NOTE, 'c'))
    st['k'] = ParagraphStyle('k', fontName=FN, fontSize=6.3, leading=8, alignment=1)
    st['kh'] = ParagraphStyle('kh', fontName=FB, fontSize=6.3, leading=8, alignment=1, textColor=colors.white)
    kr = kosha_rows(A)
    kdata = [[P(h, 'kh') for h in KOSHA_HEAD]] + [[P(v, 'k') for v in row] for row in kr]
    kw = [3, 7, 5, 5, 6, 6.5, 10, 8, 6, 8.5, 8.5, 8.5, 8.5, 8.5, 8.5, 9]
    kt = Table(kdata, colWidths=[W * w / sum(kw) for w in kw], repeatRows=1)
    ksty = [('GRID', (0, 0), (-1, -1), 0.3, colors.HexColor('#A0A7B4')), ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1D2330')),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'), ('TOPPADDING', (0, 0), (-1, -1), 1), ('BOTTOMPADDING', (0, 0), (-1, -1), 1),
            ('LEFTPADDING', (0, 0), (-1, -1), 1), ('RIGHTPADDING', (0, 0), (-1, -1), 1)]
    for ri, row in enumerate(kr, 1):
        for ci in range(9, 16):
            if row[ci] == '통증호소자':
                ksty.append(('BACKGROUND', (ci, ri), (ci, ri), colors.HexColor('#FDE8E8')))
            elif row[ci] == '관리대상자':
                ksty.append(('BACKGROUND', (ci, ri), (ci, ri), colors.HexColor('#FEF0C7')))
    kt.setStyle(TableStyle(ksty))
    el.append(kt)
    el.append(P('4-7. 통증호소자 명단 (대외비)', 'h2'))
    ml = [[x.name, x.dept, x.job_label, x.age, ', '.join(S.label('part', p) for p, v in x.part_levels.items() if v == 'pain'), x.followup or '미조치']
          for x in A['pain_list']]
    el.append(tbl(['성명', '부서', '작업', '연령', '해당 부위', '사후관리'], ml or [['해당자 없음', '', '', '', '', '']], [14, 16, 22, 8, 22, 18]))
    el.append(P('4-8. 관리대상자 명단 (통증호소자 제외)', 'h2'))
    nl = [[x.name, x.dept, x.job_label, x.age, ', '.join(S.label('part', p) for p, v in x.part_levels.items() if v == 'manage'), x.followup or '']
          for x in A['manage_list']]
    el.append(tbl(['성명', '부서', '작업', '연령', '해당 부위', '사후관리'], nl or [['해당자 없음', '', '', '', '', '']], [14, 16, 22, 8, 22, 18]))

    # 5. 종합판정
    basis, fit, pers = judge_rows(A)
    newpage()
    el += [P('5. 유해요인조사표 × 증상조사표 종합판정', 'h1'), P('5-1. 판정 기준', 'h2'), kv(JUDGE_BASIS), P('5-2. 작업별 판정 및 근거', 'h2')]
    bf = {i: ('#FDE8E8' if row[0].startswith('①') else '#FEF0C7' if row[0].startswith('②') else '#FFFFFF') for i, row in enumerate(basis)}
    el.append(tbl(['판정', '작업', '부담작업', 'A×B', '응답', '관리대상↑', '통증', '근거'], basis or [['-'] * 8],
                  [12, 18, 8, 6, 6, 10, 6, 34], center_from=9, row_fills=bf))
    el += [P('5-3. 신체부위 부합성', 'h2'),
           P('● 부담작업 관련 부위 · 숫자 = 관리대상자/통증호소자 · ◎ 부합(작업 관련성 근거) · △ 조사표에 없는 부위의 증상', 'c')]
    el.append(tbl(['작업', '부담작업'] + [S.label('part', q) for q in S.PARTS], [r_[:-1] for r_ in fit] or [['-'] * 8],
                  [22, 10, 11, 11, 11, 12, 11, 11], center_from=1))
    el += [P('5-4. 개인요인 확인 (관리대상자·통증호소자)', 'h2'),
           tbl(['성명', '판정', '작업', '증상 부위', '작업 관련', '개인요인', '판단', '사후관리'], pers or [['해당자 없음'] + [''] * 7],
               [9, 9, 14, 14, 12, 20, 12, 10])]

    # 6. 개선계획서
    newpage()
    el += [P('6. 작업환경 개선계획서', 'h1')]
    smap = {j['job'].id: j for j in A['job_summary']}
    ir = []
    for m in A['improvements']:
        js = smap.get(m.job_id, {})
        ir.append([m.priority, m.job.name if m.job else '', m.problem, m.opinion, m.plan, m.schedule,
                   f'{m.cost:,}' if m.cost else '', js.get('score') or '', js.get('symptom', ''), m.status])
    el.append(tbl(['순위', '작업명', '문제점(원인)', '근로자 의견', '개선방안', '일정', '비용(천원)', '총점', '증상', '상태'],
                  ir or [['', '등록된 개선계획 없음'] + [''] * 8], [5, 11, 20, 15, 20, 9, 8, 5, 5, 6], center_from=5))
    conf = Table([[P('위 작업환경 개선계획을 확인합니다.', 'c'), P('사업주', 'cc'), '', P('근로자 대표', 'cc'), '']],
                 colWidths=[W * 0.40, W * 0.12, W * 0.18, W * 0.12, W * 0.18], rowHeights=[12 * mm])
    conf.setStyle(TableStyle([('GRID', (1, 0), (-1, -1), 0.5, colors.black), ('BACKGROUND', (1, 0), (1, 0), colors.HexColor('#EEF1F4')),
                              ('BACKGROUND', (3, 0), (3, 0), colors.HexColor('#EEF1F4')), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE')]))
    el += [Spacer(1, 3 * mm), conf]
    if A['ranked']:
        el.append(P('개선 우선순위 선정 결과(자동)', 'h2'))
        el.append(tbl(['순위', '부서', '작업명', '부담작업(호)', 'A×B 최고', '관리대상', '통증호소'],
                      [[j['rank'], j['job'].dept, j['job'].name, ', '.join(map(str, j['items'])) or '비해당', j['score'], j['manage'], j['pain']] for j in A['ranked']],
                      [8, 18, 24, 18, 10, 10, 10], center_from=3))

    # 6. 종합의견
    el += [Spacer(1, 6 * mm), P('7. 종합 의견', 'h1')]
    for i, ptxt in enumerate(conclusion_text(A), 1):
        el += [P(f'{i}. {ptxt}', 'b'), Spacer(1, 2 * mm)]

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=15 * mm, rightMargin=15 * mm, topMargin=15 * mm, bottomMargin=18 * mm,
                            title=r.title if r else '근골격계 유해요인조사 결과보고서', author=company)
    doc.build(el, onFirstPage=lambda c, d: None, onLaterPages=footer)
    return buf.getvalue()
