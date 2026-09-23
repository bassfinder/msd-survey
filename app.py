# -*- coding: utf-8 -*-
"""근골격계 증상조사 · 유해요인조사 정기보고서 앱 (Soudal Korea EHS)

직원: QR → /survey (7개 언어) 로 증상조사표 제출 → 자유기재는 한국어 자동번역 병행저장
관리자: /admin 대시보드(부위별·부서별·작업별 집계, 관리대상/통증호소자 판정)
        작업 마스터 → 작업별 유해요인조사표(별지 제1호, 부담작업 체크 포함) → 개선계획서
        → 정기보고서 Excel / PDF 자동 생성
"""
from flask import Response as HttpResponse
from flask import Flask, render_template, request, redirect, url_for, session, abort, flash
from functools import wraps
from flask_sqlalchemy import SQLAlchemy
from datetime import datetime, timedelta, date
from urllib.parse import quote
from werkzeug.middleware.proxy_fix import ProxyFix
from deep_translator import GoogleTranslator
import os, io, json

import survey_def as S

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
app = Flask(__name__, template_folder=os.path.join(BASE_DIR, 'templates'),
            static_folder=os.path.join(BASE_DIR, 'static'))
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'msd-survey-2026')
app.config['SESSION_COOKIE_SAMESITE'] = 'Lax'
app.config['SESSION_COOKIE_HTTPONLY'] = True
_on_cloud = bool(os.environ.get('RAILWAY_PUBLIC_DOMAIN') or os.environ.get('RAILWAY_ENVIRONMENT') or os.environ.get('PUBLIC_URL')
                 or os.environ.get('DATABASE_URL', '').startswith(('postgres://', 'postgresql://')))
app.config['SESSION_COOKIE_SECURE'] = _on_cloud
app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(days=30)
app.config['MAX_CONTENT_LENGTH'] = 40 * 1024 * 1024
app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1)

_db_url = os.environ.get('DATABASE_URL', '')
if _db_url.startswith('postgres://'):
    _db_url = _db_url.replace('postgres://', 'postgresql://', 1)
# SQLITE_PATH: 로컬 테스트용 DB 파일 지정 (없으면 앱 폴더의 msd.db)
app.config['SQLALCHEMY_DATABASE_URI'] = _db_url or 'sqlite:///' + os.environ.get('SQLITE_PATH', os.path.join(BASE_DIR, 'msd.db')).replace(os.sep, '/')
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
db = SQLAlchemy(app)

ADMIN_PASSWORD = os.environ.get('ADMIN_PASSWORD', 'msd2026')
# 클라우드에서 ADMIN_PASSWORD 환경변수를 안 넣으면 기본 암호로는 로그인 못 하게 막음 (건강정보 보호)
PW_UNSET_ON_CLOUD = _on_cloud and not os.environ.get('ADMIN_PASSWORD')
# 안전의 날: 수달세이프 안전 설문과 연결 (무기명 설문과 DB는 분리, 링크만 연결). 비우면 버튼 안 보임
SAFETY_SURVEY_URL = os.environ.get('SAFETY_SURVEY_URL', '')
# 직원 공개 스위치: '0'이면 직원(비로그인)에게 설문을 열지 않음 — 관리자는 로그인 후 테스트 가능 (수달세이프 SURVEY_OPEN 과 같은 방식)
SURVEY_OPEN = os.environ.get('SURVEY_OPEN', '1') != '0'
COMPANY = os.environ.get('COMPANY_NAME', '수달코리아(주)')
SITE = os.environ.get('SITE_NAME', '안성공장')
KST = timedelta(hours=9)


def kst(dt):
    return (dt + KST) if dt else None


def jload(s, default=None):
    try:
        return json.loads(s) if s else (default if default is not None else {})
    except Exception:
        return default if default is not None else {}


def jdump(o):
    return json.dumps(o, ensure_ascii=False)


# ── 다국어 ────────────────────────────────────────────────────────────
_TEXTS = {}


_TEXTS_MTIME = None


def load_texts():
    global _TEXTS, _TEXTS_MTIME
    p = os.path.join(BASE_DIR, 'i18n', 'survey_texts.json')
    try:
        _TEXTS_MTIME = os.path.getmtime(p)
        with open(p, encoding='utf-8') as f:
            _TEXTS = json.load(f)
    except Exception:
        _TEXTS = {}


def texts_fresh():
    """번역 파일이 바뀌었으면 다시 읽기 (서버 재시작 없이 번역 수정 반영)"""
    try:
        if os.path.getmtime(os.path.join(BASE_DIR, 'i18n', 'survey_texts.json')) != _TEXTS_MTIME:
            load_texts()
    except Exception:
        pass


load_texts()


class T(dict):
    """템플릿용: t.key / t['key'] — 번역 없으면 한국어로 폴백"""
    def __init__(self, lang):
        super().__init__()
        self.lang = lang

    def __missing__(self, k):
        return _TEXTS.get(self.lang, {}).get(k) or S.KO.get(k, k)

    __getattr__ = dict.__getitem__


def translate_units(units):
    """단위작업 이름·설명을 언어별로 묶어서 번역(언어당 1회 호출). 구글 차단 시 있는 만큼만 반환"""
    import re
    out = {}
    lines = []
    for u in units:
        lines += [u.get('name', ''), u.get('desc', '')]
    src = '\n'.join(f'#{i}# {t}' for i, t in enumerate(lines) if t)
    if not src:
        return out
    for lg in S.LANGS:
        if lg == 'ko':
            continue
        try:
            res = GoogleTranslator(source='ko', target=lg).translate(src) or ''
        except Exception:
            break
        got = {int(m.group(1)): m.group(2).strip()
               for m in re.finditer(r'#\s*(\d+)\s*#\s*(.+?)(?=\s*#\s*\d+\s*#|\Z)', res, re.S)}
        out[lg] = [{'name': got.get(2 * k, ''), 'desc': got.get(2 * k + 1, '')} for k in range(len(units))]
    return out


def unit_stats(round_id, job_id):
    """조사표용: 단위작업별 근로자 응답 수·평균 작업부하/빈도 + 부담작업 호 후보 + 가장 힘든 점(원인분석 초안)"""
    agg = {}
    for r in Response.query.filter_by(round_id=round_id, job_id=job_id):
        for u in jload(r.units, []):
            if not (u.get('load') and u.get('freq')):
                continue
            g = agg.setdefault(u['name'], {'ab': [], 'sug': {}, 'hz': {}, 'txt': []})
            g['ab'].append((u['load'], u['freq']))
            for i in (u.get('sug') if u.get('sug') is not None else S.suggest_items(u)):
                g['sug'][i] = g['sug'].get(i, 0) + 1
            if u.get('hz'):
                g['hz'][u['hz']] = g['hz'].get(u['hz'], 0) + 1
            if u.get('hzt_ko') or u.get('hzt'):
                g['txt'].append(u.get('hzt_ko') or u.get('hzt'))
    out = {}
    for k, g in agg.items():
        v = g['ab']
        n = len(v)
        out[k] = {'n': n, 'a': round(sum(x for x, _ in v) / n, 1), 'b': round(sum(y for _, y in v) / n, 1),
                  'score': round(sum(x * y for x, y in v) / n, 1),
                  'sug': sorted(g['sug'].items()), 'sug_major': sorted(i for i, c in g['sug'].items() if c * 2 >= n),
                  'hz': sorted(g['hz'].items(), key=lambda kv: -kv[1]), 'txt': g['txt'][:5]}
    return out


def translate_to_korean(text, source_lang):
    if source_lang == 'ko' or not text or not text.strip():
        return text
    try:
        return GoogleTranslator(source=source_lang, target='ko').translate(text) or text
    except Exception:
        return text


# ── 모델 ──────────────────────────────────────────────────────────────
class Round(db.Model):
    """조사 회차 (예: 2026년 정기 유해요인조사)"""
    id = db.Column(db.Integer, primary_key=True)
    title = db.Column(db.String(200), nullable=False)
    kind = db.Column(db.String(50), default='정기조사')
    start_date = db.Column(db.Date)
    end_date = db.Column(db.Date)
    investigator = db.Column(db.String(100))
    agency = db.Column(db.String(100))            # 자체 / 위탁기관명
    is_open = db.Column(db.Boolean, default=False)  # 직원 설문 접수중
    overview = db.Column(db.Text)                  # 보고서 '조사 개요' 추가 서술
    conclusion = db.Column(db.Text)                # 보고서 '종합 의견'
    created_at = db.Column(db.DateTime, default=datetime.utcnow)


class Job(db.Model):
    """작업 마스터 (부서 > 공정 > 작업)"""
    id = db.Column(db.Integer, primary_key=True)
    dept = db.Column(db.String(100), nullable=False)
    process = db.Column(db.String(100))
    name = db.Column(db.String(150), nullable=False)
    workers = db.Column(db.Integer)               # 작업자 수
    description = db.Column(db.Text)
    active = db.Column(db.Boolean, default=True)
    sort = db.Column(db.Integer, default=0)
    units = db.Column(db.Text)       # 단위작업 프리셋 [{"name","loc","desc","items":[호]}] — 설문 선택지·조사표 2단계 초기값
    units_tr = db.Column(db.Text)    # {"en":[{"name","desc"}],...} 단위작업 번역 캐시

    @property
    def unit_list(self):
        return jload(self.units, [])

    @property
    def label(self):
        return f'{self.dept} / {self.name}'


class Assessment(db.Model):
    """회차별·작업별 유해요인조사표 [별지 제1호서식]"""
    id = db.Column(db.Integer, primary_key=True)
    round_id = db.Column(db.Integer, db.ForeignKey('round.id'), nullable=False)
    job_id = db.Column(db.Integer, db.ForeignKey('job.id'), nullable=False)
    survey_date = db.Column(db.Date)
    investigator = db.Column(db.String(100))
    site_status = db.Column(db.Text)     # {"equipment":{"v":"변화 없음","since":""}, "volume":..,"speed":..,"work":..}
    job_title = db.Column(db.String(150))
    tasks = db.Column(db.Text)           # [{"name","items":[호],"load","freq","workers"}]  단위작업별 부담작업(호)·작업부하A·작업빈도B·근로자수
    hazards = db.Column(db.Text)         # [{"task","items","type","cause","note"}]
    checklist = db.Column(db.Text)       # [{"unit":"단위작업명","items":[1,3]}]  ← tasks 에서 자동 생성
    note = db.Column(db.Text)
    worker_name = db.Column(db.String(100))   # 3단계 근로자명
    ergo = db.Column(db.Text)                 # 3-1 [{"unit","tool","result","judge"}]
    summary = db.Column(db.Text)              # 라. 유해요인 조사 결과(결과요약)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    job = db.relationship('Job')

    @property
    def task_list(self):
        return jload(self.tasks, [])

    @property
    def max_score(self):
        return max([int(t.get('load') or 0) * int(t.get('freq') or 0) for t in self.task_list] or [0])

    @property
    def burden_items(self):
        s = set()
        for u in jload(self.checklist, []):
            s.update(int(i) for i in u.get('items', []))
        return sorted(s)


class Photo(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    assessment_id = db.Column(db.Integer, db.ForeignKey('assessment.id'), nullable=False)
    caption = db.Column(db.String(200))
    data = db.Column(db.LargeBinary)
    mime = db.Column(db.String(50))


class Improvement(db.Model):
    """작업환경 개선계획서"""
    id = db.Column(db.Integer, primary_key=True)
    round_id = db.Column(db.Integer, db.ForeignKey('round.id'), nullable=False)
    job_id = db.Column(db.Integer, db.ForeignKey('job.id'))
    problem = db.Column(db.Text)          # 문제점(유해요인의 원인)
    opinion = db.Column(db.Text)          # 근로자 의견
    plan = db.Column(db.Text)             # 개선방안
    schedule = db.Column(db.String(100))  # 추진일정
    cost = db.Column(db.Integer)          # 천원
    priority = db.Column(db.Integer)
    status = db.Column(db.String(20), default='계획')   # 계획|진행|완료
    done_date = db.Column(db.Date)
    result = db.Column(db.Text)           # 개선 결과/효과
    job = db.relationship('Job')


class Employee(db.Model):
    """직원명단 (유해요인조사표 3단계 '근로자명' 선택용). 엑셀 명단을 불러와 부서→작업으로 연결"""
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    dept = db.Column(db.String(100))          # 명단 원래 부서명
    position = db.Column(db.String(50))
    nationality = db.Column(db.String(50))
    note = db.Column(db.String(100))          # 명단 비고 (육아휴직·출산휴가 등)
    job_id = db.Column(db.Integer, db.ForeignKey('job.id'))
    job = db.relationship('Job')

    @property
    def on_leave(self):
        return any(k in (self.note or '') for k in ('휴직', '휴가'))


class Response(db.Model):
    """근골격계질환 증상조사표 응답"""
    id = db.Column(db.Integer, primary_key=True)
    round_id = db.Column(db.Integer, db.ForeignKey('round.id'), nullable=False)
    lang = db.Column(db.String(5), default='ko')
    name = db.Column(db.String(100))
    emp_no = db.Column(db.String(50))    # (사용 안 함)
    employee_id = db.Column(db.Integer)  # 직원명단 연결 (성명 선택 시)
    age = db.Column(db.Integer)
    sex = db.Column(db.String(1))
    career_y = db.Column(db.Integer)
    career_m = db.Column(db.Integer)
    dept = db.Column(db.String(100))
    job_id = db.Column(db.Integer, db.ForeignKey('job.id'))
    job_other = db.Column(db.String(200))
    job_other_ko = db.Column(db.String(200))
    married = db.Column(db.String(1))
    cur_task = db.Column(db.Text)
    cur_task_ko = db.Column(db.Text)
    cur_y = db.Column(db.Integer)
    cur_m = db.Column(db.Integer)
    work_hours = db.Column(db.Float)
    rest_min = db.Column(db.Integer)
    rest_times = db.Column(db.Integer)
    prev_task = db.Column(db.Text)
    prev_task_ko = db.Column(db.Text)
    prev_y = db.Column(db.Integer)
    prev_m = db.Column(db.Integer)
    hobbies = db.Column(db.Text)          # ["computer",...]
    housework = db.Column(db.String(5))
    disease_yn = db.Column(db.String(1))
    diseases = db.Column(db.Text)
    disease_status = db.Column(db.String(10))
    injury_yn = db.Column(db.String(1))
    injuries = db.Column(db.Text)
    injury_status = db.Column(db.String(10))
    burden = db.Column(db.String(5))
    has_pain = db.Column(db.String(1))
    pain = db.Column(db.Text)             # {"neck":{"side","duration","intensity","frequency","lastweek","conseq":[],"other","other_ko"}}
    opinion = db.Column(db.Text)
    opinion_ko = db.Column(db.Text)
    units = db.Column(db.Text)            # 내가 하는 작업 [{"name","load","freq"}] (유해요인조사표 2단계 근로자 면담)
    top_level = db.Column(db.String(10))  # pain|manage|complaint|None
    followup = db.Column(db.String(20))   # 통증호소자 사후관리: 미조치|상담|의료기관 권고|작업전환|완료
    admin_note = db.Column(db.Text)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    job = db.relationship('Job')

    @property
    def pain_dict(self):
        return jload(self.pain, {})

    @property
    def part_levels(self):
        return S.judge_person(self.pain_dict)[0]

    @property
    def job_label(self):
        if self.job:
            return self.job.name
        return self.job_other_ko or self.job_other or ''


def ensure_columns():
    wanted = [
        # 새 컬럼 추가 시 ('table','col','TYPE') 한 줄씩
        ('response', 'followup', 'VARCHAR(20)'),
        ('response', 'injury_status', 'VARCHAR(10)'),
        ('assessment', 'worker_name', 'VARCHAR(100)'),
        ('assessment', 'ergo', 'TEXT'),
        ('assessment', 'summary', 'TEXT'),
        ('employee', 'note', 'VARCHAR(100)'),
        ('job', 'units', 'TEXT'),
        ('response', 'employee_id', 'INTEGER'),
        ('job', 'units_tr', 'TEXT'),
        ('response', 'units', 'TEXT'),
    ]
    is_pg = str(db.engine.url).startswith('postgresql')
    with db.engine.connect() as conn:
        for table, col, coltype in wanted:
            try:
                if is_pg:
                    conn.execute(db.text(f'ALTER TABLE "{table}" ADD COLUMN IF NOT EXISTS {col} {coltype}'))
                else:
                    cols = [r[1] for r in conn.execute(db.text(f'PRAGMA table_info({table})'))]
                    if col not in cols:
                        conn.execute(db.text(f'ALTER TABLE {table} ADD COLUMN {col} {coltype}'))
                conn.commit()
            except Exception as e:
                print(f'[migrate skip] {table}.{col}: {e}')


with app.app_context():
    db.create_all()
    ensure_columns()


# ── 공통 헬퍼 ─────────────────────────────────────────────────────────
def to_int(v):
    try:
        return int(str(v).strip())
    except Exception:
        return None


def to_float(v):
    try:
        return float(str(v).strip())
    except Exception:
        return None


def to_date(v):
    try:
        return datetime.strptime(v, '%Y-%m-%d').date()
    except Exception:
        return None


def open_round():
    return Round.query.filter_by(is_open=True).order_by(Round.id.desc()).first()


def current_round():
    rid = session.get('round_id')
    r = db.session.get(Round, rid) if rid else None
    return r or Round.query.order_by(Round.id.desc()).first()


def admin_required(f):
    @wraps(f)
    def deco(*a, **kw):
        if not session.get('admin'):
            return redirect(url_for('admin_login', next=request.path))
        return f(*a, **kw)
    return deco


@app.context_processor
def inject():
    ctx = {'COMPANY': COMPANY, 'SITE': SITE, 'S': S, 'kst': kst, 'survey_public': SURVEY_OPEN}
    if session.get('admin'):
        ctx['cur_round'] = current_round()
        ctx['all_rounds'] = Round.query.order_by(Round.id.desc()).all()
    return ctx


# ── 직원용 증상조사표 ─────────────────────────────────────────────────
@app.route('/healthz')
def healthz():
    # 배포 점검용 — 비밀값은 내보내지 않고 '설정됨' 여부만
    return {'ok': True, 'db': db.engine.url.get_backend_name(), 'admin_password_set': bool(os.environ.get('ADMIN_PASSWORD')),
            'secret_key_set': bool(os.environ.get('SECRET_KEY')), 'survey_open': SURVEY_OPEN, 'cloud': _on_cloud,
            'rounds': Round.query.count(), 'open_round': bool(open_round()), 'jobs': Job.query.count(),
            'employees': Employee.query.count(), 'responses': Response.query.count()}


@app.route('/')
def home():
    return redirect(url_for('survey'))


@app.route('/survey', methods=['GET', 'POST'])
def survey():
    lang = request.values.get('lang') or session.get('lang') or 'ko'
    if lang not in S.LANGS:
        lang = 'ko'
    session['lang'] = lang
    texts_fresh()
    t = T(lang)
    rnd = open_round() if (SURVEY_OPEN or session.get('admin')) else None
    # 관리자 미리보기: 접수 전이라도 설문 화면을 볼 수 있음(제출은 저장 안 함)
    preview = bool(session.get('admin') and (request.values.get('preview') or not rnd))
    if not rnd and preview:
        rnd = current_round() or Round(title='(회차 없음 — 미리보기)')
    if not rnd:
        return render_template('survey_closed.html', t=t, lang=lang)
    jobs = Job.query.filter_by(active=True).order_by(Job.sort, Job.dept, Job.name).all()
    depts = list(dict.fromkeys(j.dept for j in jobs))   # 작업 마스터 정렬순서대로
    error = None
    if request.method == 'POST' and preview:
        flash('미리보기에서는 제출이 저장되지 않습니다. 실제 접수는 [조사회차 → 접수 시작] 후 진행하세요.')
        return redirect(url_for('survey', preview=1, lang=lang))
    if request.method == 'POST':
        f = request.form
        pain = {}
        has_pain = f.get('has_pain')
        if has_pain == 'Y':
            for part in S.PARTS:
                if f.get(f'part_{part}'):
                    d = {
                        'side': f.get(f'{part}_side') if part in S.SIDED_PARTS else None,
                        'duration': f.get(f'{part}_duration'),
                        'intensity': f.get(f'{part}_intensity'),
                        'frequency': f.get(f'{part}_frequency'),
                        'lastweek': f.get(f'{part}_lastweek'),
                        'conseq': f.getlist(f'{part}_conseq'),
                        'other': (f.get(f'{part}_other') or '').strip(),
                    }
                    pain[part] = d
        emp = db.session.get(Employee, to_int(f.get('employee_id'))) if to_int(f.get('employee_id')) else None
        f = f.copy()
        if emp:
            f['name'] = emp.name
            if not f.get('dept') and emp.job:
                f['dept'] = emp.job.dept
            if not f.get('job_id') and emp.job_id:
                f['job_id'] = str(emp.job_id)
        if not f.get('consent'):
            error = t.err_consent
        elif not all([f.get('name', '').strip(), f.get('age'), f.get('sex'), f.get('dept'), has_pain]):
            error = t.err_required
        elif has_pain == 'Y' and not pain:
            error = t.err_part
        elif any(not (p['duration'] and p['intensity'] and p['frequency']) for p in pain.values()):
            error = t.err_required
        if not error:
            for p in pain.values():
                if p['other']:
                    p['other_ko'] = translate_to_korean(p['other'], lang)
            job_id = to_int(f.get('job_id'))
            r = Response(
                round_id=rnd.id, lang=lang,
                name=f.get('name', '').strip(), employee_id=emp.id if emp else None,
                age=to_int(f.get('age')), sex=f.get('sex'),
                career_y=to_int(f.get('career_y')), career_m=to_int(f.get('career_m')),
                dept=f.get('dept'), job_id=job_id if job_id else None,
                job_other=f.get('job_other', '').strip(),
                cur_task=f.get('cur_task', '').strip(),
                cur_y=to_int(f.get('cur_y')), cur_m=to_int(f.get('cur_m')),
                work_hours=to_float(f.get('work_hours')),
                rest_min=to_int(f.get('rest_min')), rest_times=to_int(f.get('rest_times')),
                prev_task=f.get('prev_task', '').strip(),
                prev_y=to_int(f.get('prev_y')), prev_m=to_int(f.get('prev_m')),
                hobbies=jdump(f.getlist('hobby')), housework=f.get('housework'),
                disease_yn=f.get('disease_yn'), diseases=jdump(f.getlist('disease')),
                disease_status=f.get('disease_status'),
                injury_yn=f.get('injury_yn'), injuries=jdump(f.getlist('injury')),
                injury_status=f.get('injury_status'),
                married=f.get('married'),
                burden=f.get('burden'), has_pain=has_pain, pain=jdump(pain),
                opinion=f.get('opinion', '').strip(),
            )
            for fld in ('job_other', 'cur_task', 'prev_task', 'opinion'):
                setattr(r, fld + '_ko', translate_to_korean(getattr(r, fld), lang))
            job = db.session.get(Job, job_id) if job_id else None
            chosen = []
            for i, u in enumerate(job.unit_list if job else []):
                if f.get(f'u_{i}'):
                    c_ = {'name': u['name'], 'load': to_int(f.get(f'u_{i}_load')), 'freq': to_int(f.get(f'u_{i}_freq')),
                          'hrs': f.get(f'u_{i}_hrs'), 'wt': f.get(f'u_{i}_wt'), 'cnt': f.get(f'u_{i}_cnt'),
                          'pos': f.getlist(f'u_{i}_pos'), 'hz': f.get(f'u_{i}_hz'), 'hzt': (f.get(f'u_{i}_hzt') or '').strip()}
                    if c_['hzt']:
                        c_['hzt_ko'] = translate_to_korean(c_['hzt'], lang)
                    c_['sug'] = S.suggest_items(c_)
                    chosen.append(c_)
            r.units = jdump(chosen)
            if not r.cur_task and chosen:
                r.cur_task = r.cur_task_ko = ', '.join(c['name'] for c in chosen)
            r.top_level = S.judge_person(pain)[1]
            replaced = 0
            if r.employee_id:
                for old in Response.query.filter_by(round_id=rnd.id, employee_id=r.employee_id).all():
                    db.session.delete(old)
                    replaced += 1
            db.session.add(r)
            db.session.commit()
            session['last_rid'] = r.id
            session['replaced'] = replaced
            return redirect(url_for('survey_done'))
    def units_js(j):
        tr = jload(j.units_tr, {}).get(lang, [])
        return [{'i': i, 'name': u.get('name', ''), 'loc': u.get('loc', ''), 'desc': u.get('desc', ''),
                 'tname': (tr[i].get('name') if i < len(tr) else '') or '', 'tdesc': (tr[i].get('desc') if i < len(tr) else '') or ''}
                for i, u in enumerate(j.unit_list)]
    jobs_js = [{'id': j.id, 'dept': j.dept, 'name': j.name, 'units': units_js(j)} for j in jobs]
    form_units = {k: request.form.get(k) for k in request.form if k.startswith('u_')}
    order = {j.id: (j.sort or 0, i) for i, j in enumerate(jobs)}
    emps = sorted(Employee.query.filter(Employee.job_id.in_([j.id for j in jobs])).all(),
                  key=lambda e: (order.get(e.job_id, (999, 999)), e.name))
    emps_js = [{'id': e.id, 'name': e.name, 'dept': e.job.dept if e.job else '', 'job_id': e.job_id,
                'pos': e.position or '', 'leave': e.on_leave} for e in emps]
    mine = db.session.get(Response, session.get('last_rid')) if session.get('last_rid') and not preview else None
    if mine and mine.round_id != rnd.id:
        mine = None
    return render_template('survey.html', t=t, lang=lang, rnd=rnd, depts=depts, jobs_js=jobs_js, form_units=form_units, preview=preview, emps_js=emps_js, mine=mine,
                           error=error, form=request.form)


@app.route('/survey/done')
def survey_done():
    lang = session.get('lang', 'ko')
    r = db.session.get(Response, session.get('last_rid')) if session.get('last_rid') else None
    return render_template('survey_done.html', t=T(lang), lang=lang, r=r, is_admin=bool(session.get('admin')), safety_url=SAFETY_SURVEY_URL,
                           replaced=session.pop('replaced', 0), kst=kst)


@app.route('/qr')
@admin_required
def qr_page():
    url = request.url_root.rstrip('/') + url_for('survey')
    return render_template('qr.html', url=url)


@app.route('/qr.png')
@admin_required
def qr_png():
    import qrcode
    url = request.url_root.rstrip('/') + url_for('survey')
    img = qrcode.make(url, box_size=12, border=2)
    buf = io.BytesIO()
    img.save(buf, format='PNG')
    return HttpResponse(buf.getvalue(), mimetype='image/png')


# ── 관리자 인증 ───────────────────────────────────────────────────────
@app.route('/admin/login', methods=['GET', 'POST'])
def admin_login():
    error = None
    if request.method == 'POST':
        if PW_UNSET_ON_CLOUD:
            error = '서버에 ADMIN_PASSWORD 환경변수가 설정되지 않아 로그인을 막았습니다. Railway Variables에 설정하세요.'
        elif request.form.get('password') == ADMIN_PASSWORD:
            session.permanent = True
            session['admin'] = True
            return redirect(request.args.get('next') or url_for('admin'))
        else:
            error = '암호가 틀렸습니다.'
    return render_template('admin_login.html', error=error)


@app.route('/admin/logout')
def admin_logout():
    session.clear()
    return redirect(url_for('admin_login'))


@app.route('/admin/round/select/<int:rid>')
@admin_required
def round_select(rid):
    session['round_id'] = rid
    return redirect(request.referrer or url_for('admin'))


# ── 분석 ──────────────────────────────────────────────────────────────
def analyze(rnd):
    """대시보드·Excel·PDF 공통 집계"""
    res = Response.query.filter_by(round_id=rnd.id).order_by(Response.id).all() if rnd else []
    n = len(res)
    pct = lambda a, b: round(a * 100.0 / b, 1) if b else 0.0

    def age_grp(a):
        if a is None: return '미기재'
        if a < 30: return '29세 이하'
        if a < 40: return '30~39세'
        if a < 50: return '40~49세'
        if a < 60: return '50~59세'
        return '60세 이상'

    def career_grp(r):
        y = (r.career_y or 0) + (r.career_m or 0) / 12.0
        if r.career_y is None and r.career_m is None: return '미기재'
        if y < 1: return '1년 미만'
        if y < 5: return '1~5년 미만'
        if y < 10: return '5~10년 미만'
        return '10년 이상'

    def count(keyf, order=None):
        c = {}
        for r in res:
            k = keyf(r)
            c[k] = c.get(k, 0) + 1
        keys = order if order else sorted(c)
        return [(k, c.get(k, 0), pct(c.get(k, 0), n)) for k in keys if c.get(k, 0) or order]

    general = {
        '성별': count(lambda r: S.label('sex', r.sex) or '미기재', ['남', '여']),
        '연령': count(lambda r: age_grp(r.age), ['29세 이하', '30~39세', '40~49세', '50~59세', '60세 이상']),
        '현 직장경력': count(career_grp, ['1년 미만', '1~5년 미만', '5~10년 미만', '10년 이상']),
        '부서': count(lambda r: r.dept or '미기재'),
        '육체적 부담정도': count(lambda r: S.label('burden', r.burden) or '미기재',
                           [S.label('burden', b) for b in S.BURDEN]),
        '응답 언어': count(lambda r: S.LANG_NAMES.get(r.lang, r.lang)),
    }
    hobby_c = {h: sum(1 for r in res if h in jload(r.hobbies, [])) for h in S.HOBBY}
    general_extra = {
        '여가·취미활동': [(S.label('hobby', h), hobby_c[h], pct(hobby_c[h], n)) for h in S.HOBBY],
        '질병 진단(예)': [(S.label('disease', d), sum(1 for r in res if r.disease_yn == 'Y' and d in jload(r.diseases, [])), 0) for d in S.DISEASE],
        '과거 상해(예)': [(S.label('injury', d), sum(1 for r in res if r.injury_yn == 'Y' and d in jload(r.injuries, [])), 0) for d in S.INJURY_PARTS],
    }
    for k in ('질병 진단(예)', '과거 상해(예)'):
        general_extra[k] = [(a, b, pct(b, n)) for a, b, _ in general_extra[k]]

    # 부위별
    parts = []
    for part in S.PARTS:
        lv = [r.part_levels.get(part) for r in res]
        c = sum(1 for v in lv if v)
        ni = sum(1 for v in lv if v in ('manage', 'pain'))
        ma = sum(1 for v in lv if v == 'pain')
        dur = {d: sum(1 for r in res if r.pain_dict.get(part, {}).get('duration') == d) for d in S.DURATION}
        inten = {d: sum(1 for r in res if r.pain_dict.get(part, {}).get('intensity') == d) for d in S.INTENSITY}
        freq = {d: sum(1 for r in res if r.pain_dict.get(part, {}).get('frequency') == d) for d in S.FREQUENCY}
        parts.append({'part': part, 'label': S.label('part', part), 'complaint': c, 'manage': ni, 'pain': ma,
                      'complaint_p': pct(c, n), 'manage_p': pct(ni, n), 'pain_p': pct(ma, n),
                      'dur': dur, 'int': inten, 'freq': freq})
    persons = {
        'complaint': sum(1 for r in res if r.top_level),
        'manage': sum(1 for r in res if r.top_level in ('manage', 'pain')),
        'pain': sum(1 for r in res if r.top_level == 'pain'),
    }

    def group_rows(keyf):
        g = {}
        for r in res:
            g.setdefault(keyf(r), []).append(r)
        rows = []
        for k in sorted(g, key=lambda x: str(x)):
            rs = g[k]
            m = len(rs)
            c = sum(1 for r in rs if r.top_level)
            ni = sum(1 for r in rs if r.top_level in ('manage', 'pain'))
            ma = sum(1 for r in rs if r.top_level == 'pain')
            part_c = {p: sum(1 for r in rs if r.part_levels.get(p)) for p in S.PARTS}
            rows.append({'key': k, 'n': m, 'complaint': c, 'manage': ni, 'pain': ma,
                         'complaint_p': pct(c, m), 'manage_p': pct(ni, m), 'pain_p': pct(ma, m), 'parts': part_c})
        return rows

    by_dept = group_rows(lambda r: r.dept or '미기재')
    by_job = group_rows(lambda r: (r.dept or '') + ' / ' + (r.job_label or '미기재'))

    pain_list = [r for r in res if r.top_level == 'pain']
    manage_list = [r for r in res if r.top_level == 'manage']

    # 작업별 종합 (기본조사표 점수 + 체크리스트 + 증상) → 개선 우선순위 제안
    assess = Assessment.query.filter_by(round_id=rnd.id).all() if rnd else []
    amap = {a.job_id: a for a in assess}
    jobs = Job.query.order_by(Job.sort, Job.dept, Job.name).all()
    job_summary = []
    for j in jobs:
        a = amap.get(j.id)
        rs = [r for r in res if r.job_id == j.id]
        if not a and not rs and not j.active:
            continue
        m = len(rs)
        c = sum(1 for r in rs if r.top_level)
        ni = sum(1 for r in rs if r.top_level in ('manage', 'pain'))
        ma = sum(1 for r in rs if r.top_level == 'pain')
        items = a.burden_items if a else []
        score = a.max_score if a else 0
        job_summary.append({'job': j, 'assess': a, 'n': m, 'complaint': c, 'manage': ni, 'pain': ma,
                            'items': items, 'is_burden': bool(items), 'score': score,
                            'symptom': 'O' if c else 'X'})
    # ── 종합판정: 부위 부합성 · 2×2 · 개인요인 ──
    site_rate = persons['manage'] / n if n else 0          # 사업장 전체 관리대상자 이상 비율
    for x in job_summary:
        rs = [r for r in res if r.job_id == x['job'].id]
        exp = sorted({p for i in x['items'] for p in S.BURDEN_PARTS.get(i, [])}, key=S.PARTS.index)
        sym = {p: {'manage': sum(1 for r in rs if r.part_levels.get(p) == 'manage'),
                   'pain': sum(1 for r in rs if r.part_levels.get(p) == 'pain'),
                   'complaint': sum(1 for r in rs if r.part_levels.get(p))} for p in S.PARTS}
        match = [p for p in exp if sym[p]['manage'] + sym[p]['pain'] > 0]
        unexpected = [p for p in S.PARTS if p not in exp and sym[p]['manage'] + sym[p]['pain'] > 0]
        expo_high = x['is_burden'] and x['score'] >= S.EXPOSURE_SCORE_HIGH
        rate = x['manage'] / x['n'] if x['n'] else 0
        sym_high = x['n'] > 0 and (x['pain'] > 0 or (x['manage'] > 0 and rate >= site_rate))
        q = (0 if not x['n'] else 5 if not x['assess'] else
             1 if expo_high and sym_high else 2 if expo_high else 3 if sym_high else 4)
        why = []
        if x['items']:
            why.append('부담작업 ' + ','.join(map(str, x['items'])) + '호')
        if x['score']:
            why.append(f"A×B 최고 {x['score']}점" + (' (높음)' if x['score'] >= S.EXPOSURE_SCORE_HIGH else ''))
        if x['n']:
            why.append(f"관리대상자 이상 {x['manage']}/{x['n']}명({round(rate * 100)}%, 사업장 {round(site_rate * 100)}%)" + (f", 통증호소자 {x['pain']}명" if x['pain'] else ''))
        if match:
            why.append('부위 부합: ' + '·'.join(S.label('part', p) for p in match))
        x.update({'exp_parts': exp, 'sym_parts': sym, 'match': match, 'unexpected': unexpected,
                  'expo_high': expo_high, 'sym_high': sym_high, 'rate': round(rate * 100, 1),
                  'quadrant': q, 'q_label': S.QUADRANT[q][0], 'why': ' / '.join(why)})
    quad = {q: [x for x in job_summary if x.get('quadrant') == q] for q in (1, 2, 5, 3, 4, 0)}

    # 개인요인: 관리대상자·통증호소자별로 증상 부위와 겹치는 과거 상해·질병·취미·가사노동
    jmap = {x['job'].id: x for x in job_summary}
    personal = []
    for r in res:
        if r.top_level not in ('manage', 'pain'):
            continue
        hot = [p for p, v in r.part_levels.items() if v in ('manage', 'pain')]
        inj = [p for p in jload(r.injuries, []) if p in hot] if r.injury_yn == 'Y' else []
        hob = sorted({h for h in jload(r.hobbies, []) for p in S.HOBBY_PARTS.get(h, []) if p in hot})
        flags = []
        if inj:
            flags.append('과거 상해(' + '·'.join(S.label('part', p) for p in inj) + ')')
        if r.disease_yn == 'Y':
            flags.append('질병(' + ','.join(S.label('disease', d) for d in jload(r.diseases, [])) + ')')
        if hob:
            flags.append('취미(' + ','.join(S.label('hobby', h) for h in hob) + ')')
        if r.housework in ('h3', 'h4'):
            flags.append('가사노동 ' + S.label('housework', r.housework))
        jx = jmap.get(r.job_id)
        work_rel = [p for p in hot if jx and p in jx['exp_parts']]
        if jx and not jx['assess']:
            judge = '조사표 작성 후 판단'
        elif work_rel and not inj and r.disease_yn != 'Y':
            judge = '작업 관련성 높음'
        elif work_rel:
            judge = '작업+개인요인 혼재'
        elif flags:
            judge = '개인요인 가능성'
        else:
            judge = '작업 관련성 확인 필요'
        personal.append({'r': r, 'parts': hot, 'flags': flags, 'judge': judge,
                         'work_parts': work_rel})

    ranked = sorted([x for x in job_summary if x['is_burden'] or x['manage']],
                    key=lambda x: ({1: 0, 2: 1, 5: 2, 3: 3, 4: 4, 0: 5}[x['quadrant']], -x['pain'], -x['manage'],
                                   -x['score'], -(x['job'].workers or 0), -len(x['items'])))
    for i, x in enumerate(ranked, 1):
        x['rank'] = i

    unit_rows = []   # 데이터화: 응답자 × 선택 작업 1행
    for r in res:
        for u in jload(r.units, []):
            unit_rows.append([r.id, r.name, r.dept, r.job_label, u.get('name'), u.get('load'), u.get('freq'),
                              (u.get('load') or 0) * (u.get('freq') or 0) or None, S.KOSHA_KO[r.top_level],
                              S.label('u_hrs', u.get('hrs')) if u.get('hrs') else '', S.label('u_wt', u.get('wt')) if u.get('wt') else '',
                              S.label('u_cnt', u.get('cnt')) if u.get('cnt') else '', ', '.join(S.label('u_pos', x) for x in u.get('pos') or []),
                              S.U_HZ_TYPE.get(u.get('hz'), ''), u.get('hzt_ko') or u.get('hzt') or '',
                              ', '.join(map(str, u.get('sug') if u.get('sug') is not None else S.suggest_items(u))) or ''])
    ustat_by_job = {j['job'].id: unit_stats(rnd.id, j['job'].id) for j in job_summary} if rnd else {}
    for j in job_summary:
        j['ustats'] = ustat_by_job.get(j['job'].id, {})
    imps = Improvement.query.filter_by(round_id=rnd.id).order_by(Improvement.priority.is_(None), Improvement.priority, Improvement.id).all() if rnd else []
    return {'round': rnd, 'responses': res, 'n': n, 'general': general, 'general_extra': general_extra,
            'parts': parts, 'persons': persons,
            'persons_p': {k: pct(v, n) for k, v in persons.items()},
            'by_dept': by_dept, 'by_job': by_job, 'pain_list': pain_list, 'manage_list': manage_list,
            'unit_rows': unit_rows, 'ustat_by_job': ustat_by_job, 'quad': quad, 'personal': personal,
            'site_rate': round(site_rate * 100, 1),
            'job_summary': job_summary, 'ranked': ranked, 'assessments': assess, 'improvements': imps,
            'target_workers': sum(j.workers or 0 for j in jobs if j.active)}


# ── 관리자: 대시보드 ──────────────────────────────────────────────────
@app.route('/admin')
@admin_required
def admin():
    rnd = current_round()
    if not rnd:
        return redirect(url_for('rounds'))
    A = analyze(rnd)
    return render_template('admin_dashboard.html', A=A, rnd=rnd)


# ── 회차 ──────────────────────────────────────────────────────────────
@app.route('/admin/rounds', methods=['GET', 'POST'])
@admin_required
def rounds():
    if request.method == 'POST':
        f = request.form
        rid = to_int(f.get('id'))
        r = db.session.get(Round, rid) if rid else Round()
        r.title = f.get('title', '').strip() or f'{date.today().year}년 근골격계부담작업 유해요인조사'
        r.kind = f.get('kind')
        r.start_date = to_date(f.get('start_date'))
        r.end_date = to_date(f.get('end_date'))
        r.investigator = f.get('investigator', '').strip()
        r.agency = f.get('agency', '').strip()
        r.overview = f.get('overview', '').strip()
        r.conclusion = f.get('conclusion', '').strip()
        if not rid:
            db.session.add(r)
        db.session.commit()
        session['round_id'] = r.id
        flash('저장했습니다.')
        return redirect(url_for('rounds'))
    edit = db.session.get(Round, to_int(request.args.get('edit'))) if request.args.get('edit') else None
    rows = Round.query.order_by(Round.id.desc()).all()
    counts = {r.id: Response.query.filter_by(round_id=r.id).count() for r in rows}
    return render_template('admin_rounds.html', rows=rows, edit=edit, counts=counts)


@app.route('/admin/round/<int:rid>/toggle', methods=['POST'])
@admin_required
def round_toggle(rid):
    r = db.session.get(Round, rid) or abort(404)
    opening = not r.is_open
    if opening:   # 동시에 하나만 접수
        Round.query.update({Round.is_open: False})
    r.is_open = opening
    db.session.commit()
    flash(f"'{r.title}' 설문 접수를 {'시작' if opening else '마감'}했습니다.")
    return redirect(url_for('rounds'))


@app.route('/admin/round/<int:rid>/delete', methods=['POST'])
@admin_required
def round_delete(rid):
    r = db.session.get(Round, rid) or abort(404)
    if Response.query.filter_by(round_id=rid).count():
        flash('응답이 있는 회차는 삭제할 수 없습니다.')
        return redirect(url_for('rounds'))
    for a in Assessment.query.filter_by(round_id=rid):
        Photo.query.filter_by(assessment_id=a.id).delete()
        db.session.delete(a)
    Improvement.query.filter_by(round_id=rid).delete()
    db.session.delete(r)
    db.session.commit()
    session.pop('round_id', None)
    return redirect(url_for('rounds'))


# ── 작업 마스터 ───────────────────────────────────────────────────────
@app.route('/admin/jobs', methods=['GET', 'POST'])
@admin_required
def jobs():
    if request.method == 'POST':
        f = request.form
        jid = to_int(f.get('id'))
        j = db.session.get(Job, jid) if jid else Job()
        j.dept = f.get('dept', '').strip()
        j.process = f.get('process', '').strip()
        j.name = f.get('name', '').strip()
        j.workers = to_int(f.get('workers'))
        j.description = f.get('description', '').strip()
        old = {u['name']: u for u in j.unit_list}
        units = []
        for line in (f.get('units_text') or '').splitlines():
            parts = [x.strip() for x in line.split('|')]
            if parts and parts[0]:
                units.append({'name': parts[0], 'loc': parts[1] if len(parts) > 1 else '',
                              'desc': parts[2] if len(parts) > 2 else '', 'items': old.get(parts[0], {}).get('items', [])})
        if jdump(units) != (j.units or '[]'):
            j.units = jdump(units)
            j.units_tr = jdump(translate_units(units)) if units else None
        j.sort = to_int(f.get('sort')) or 0
        j.active = bool(f.get('active'))
        if j.dept and j.name:
            if not jid:
                db.session.add(j)
            db.session.commit()
            flash('저장했습니다.')
        return redirect(url_for('jobs'))
    rnd = current_round()
    rows = Job.query.order_by(Job.sort, Job.dept, Job.name).all()
    amap = {a.job_id: a for a in Assessment.query.filter_by(round_id=rnd.id)} if rnd else {}
    edit = db.session.get(Job, to_int(request.args.get('edit'))) if request.args.get('edit') else None
    return render_template('admin_jobs.html', rows=rows, amap=amap, edit=edit, rnd=rnd)


def norm_dept(d):
    """명단 부서명 → 작업 마스터 부서명 (물류팀[출고] → 물류팀(출고), 물류팀[관리,내부] → 물류팀(관리))"""
    d = (d or '').strip().replace('[', '(').replace(']', ')')
    return d.replace('(관리,내부)', '(관리)')


def parse_roster(fileobj):
    """직원명단 엑셀: '성명'·'부서' 머리글이 있는 첫 시트에서 재직자(퇴사일 빈칸)만 읽음"""
    import openpyxl
    wb = openpyxl.load_workbook(fileobj, data_only=True, read_only=True)
    for ws in wb.worksheets:
        head, out = None, []
        for row in ws.iter_rows(max_col=40, values_only=True):
            cells = [str(c).strip() if c is not None else '' for c in row]
            if head is None:
                if '성명' in cells and '부서' in cells:
                    head = {h: i for i, h in enumerate(cells) if h}
                continue
            g = lambda k: row[head[k]] if k in head and head[k] < len(row) else None
            name, dept, nat = g('성명'), g('부서'), g('국적')
            if not isinstance(name, str) or not name.strip() or not isinstance(dept, str):
                continue
            if g('퇴사일'):
                continue
            if '국적' in head and not isinstance(nat, str):     # 숫자만 있는 집계행 제외
                continue
            out.append({'name': name.strip(), 'dept': dept.strip(), 'position': str(g('직책') or '').strip(),
                        'nationality': (nat or '').strip(), 'note': str(g('비고') or '').strip()})
        if head:
            return out
    return []


MASTER_VERSION = 1


@app.route('/admin/master.json')
@admin_required
def master_export():
    """기본 데이터 내보내기: 작업(단위작업·번역 포함) + 직원명단 → 서버로 옮길 때 사용"""
    data = {'version': MASTER_VERSION, 'exported_at': kst(datetime.utcnow()).strftime('%Y-%m-%d %H:%M'),
            'jobs': [{'key': j.id, 'dept': j.dept, 'process': j.process, 'name': j.name, 'workers': j.workers,
                      'description': j.description, 'active': j.active, 'sort': j.sort, 'units': j.units, 'units_tr': j.units_tr}
                     for j in Job.query.order_by(Job.sort, Job.id)],
            'employees': [{'name': e.name, 'dept': e.dept, 'position': e.position, 'nationality': e.nationality,
                           'note': e.note, 'job_key': e.job_id} for e in Employee.query.order_by(Employee.id)]}
    stamp = kst(datetime.utcnow()).strftime('%Y%m%d')
    return _download(json.dumps(data, ensure_ascii=False, indent=1).encode('utf-8'), f'근골격계_기본데이터_{stamp}.json', 'application/json')


@app.route('/admin/master-import', methods=['POST'])
@admin_required
def master_import():
    """기본 데이터 가져오기: 같은 부서+작업명은 갱신, 없으면 추가. 직원명단은 통째로 교체"""
    up = request.files.get('master')
    try:
        data = json.loads(up.read().decode('utf-8'))
        assert data.get('version') == MASTER_VERSION
    except Exception:
        flash('파일을 읽지 못했습니다. [기본 데이터 내보내기]로 받은 JSON 파일인지 확인하세요.')
        return redirect(url_for('jobs'))
    keymap, added, updated = {}, 0, 0
    for d in data.get('jobs', []):
        j = Job.query.filter_by(dept=d['dept'], name=d['name']).first()
        if j:
            updated += 1
        else:
            j = Job(dept=d['dept'], name=d['name'])
            db.session.add(j)
            added += 1
        for k in ('process', 'workers', 'description', 'active', 'sort', 'units', 'units_tr'):
            setattr(j, k, d.get(k))
        db.session.flush()
        keymap[d['key']] = j.id
    emps = data.get('employees', [])
    if emps:
        Employee.query.delete()
        for e in emps:
            db.session.add(Employee(name=e['name'], dept=e.get('dept'), position=e.get('position'), nationality=e.get('nationality'),
                                    note=e.get('note'), job_id=keymap.get(e.get('job_key'))))
    db.session.commit()
    flash(f"기본 데이터를 가져왔습니다 — 작업 추가 {added}·갱신 {updated}, 직원명단 {len(emps)}명 (내보낸 시각 {data.get('exported_at')})")
    return redirect(url_for('jobs'))


@app.route('/admin/employees', methods=['GET', 'POST'])
@admin_required
def employees():
    jobs_all = Job.query.order_by(Job.sort, Job.dept, Job.name).all()
    if request.method == 'POST':
        up = request.files.get('roster')
        if up and up.filename:
            rows = parse_roster(io.BytesIO(up.read()))
            if not rows:
                flash("명단을 읽지 못했습니다. '성명'·'부서' 머리글이 있는 엑셀인지 확인하세요.")
                return redirect(url_for('employees'))
            by_dept = {}
            for j in jobs_all:
                by_dept.setdefault(j.dept, j)           # 부서의 첫 작업에 연결 (필요하면 화면에서 변경)
            old = {e.name: e.job_id for e in Employee.query}
            Employee.query.delete()
            miss = set()
            for r in rows:
                j = by_dept.get(norm_dept(r['dept']))
                if not j:
                    miss.add(r['dept'])
                db.session.add(Employee(name=r['name'], dept=r['dept'], position=r['position'], nationality=r['nationality'], note=r['note'],
                                        job_id=old.get(r['name']) or (j.id if j else None)))
            db.session.commit()
            flash(f"직원 {len(rows)}명을 불러왔습니다." + (f" 작업과 연결 안 된 부서: {', '.join(sorted(miss))}" if miss else ''))
        else:                                            # 작업 연결 변경
            for e in Employee.query:
                v = request.form.get(f'job_{e.id}')
                if v is not None:
                    e.job_id = to_int(v)
            db.session.commit()
            flash('작업 연결을 저장했습니다.')
        return redirect(url_for('employees'))
    emps = Employee.query.order_by(Employee.dept, Employee.name).all()
    return render_template('admin_employees.html', emps=emps, jobs_all=jobs_all)


@app.route('/admin/job/<int:jid>/delete', methods=['POST'])
@admin_required
def job_delete(jid):
    j = db.session.get(Job, jid) or abort(404)
    if Response.query.filter_by(job_id=jid).count() or Assessment.query.filter_by(job_id=jid).count():
        j.active = False      # 기록이 있으면 비활성화만
        flash('조사 기록이 있어 삭제 대신 비활성화했습니다.')
    else:
        Employee.query.filter_by(job_id=jid).update({Employee.job_id: None})
        db.session.delete(j)
    db.session.commit()
    return redirect(url_for('jobs'))


# ── 유해요인조사표 [별지 제1호서식] ───────────────────────────────────────────
@app.route('/admin/assess/<int:jid>', methods=['GET', 'POST'])
@admin_required
def assess(jid):
    rnd = current_round() or abort(400)
    job = db.session.get(Job, jid) or abort(404)
    a = Assessment.query.filter_by(round_id=rnd.id, job_id=jid).first()
    if request.method == 'POST':
        f = request.form
        if not a:
            a = Assessment(round_id=rnd.id, job_id=jid)
            db.session.add(a)
        a.survey_date = to_date(f.get('survey_date'))
        a.investigator = f.get('investigator', '').strip()
        a.job_title = f.get('job_title', '').strip()
        names = f.getlist('worker_pick') + [x.strip() for x in (f.get('worker_extra') or '').split(',')]
        a.worker_name = ', '.join(dict.fromkeys(n for n in names if n))
        a.site_status = jdump({k: {'v': f.get(f'st_{k}') or '변화 없음', 'since': f.get(f'st_{k}_since', '').strip()}
                               for k, _ in S.SITE_ITEMS})
        # 2단계: 단위작업명 | 부담작업(호) | 작업부하(A) | 작업빈도(B)  — 카드 순번(t_i)은 저장 직전 JS가 재정렬
        tasks = []
        for i, nm in enumerate(f.getlist('t_name')):
            if nm.strip():
                tasks.append({'name': nm.strip(), 'items': [n for n in range(1, 12) if f.get(f't_{i}_{n}')],
                              'load': to_int(f.getlist('t_load')[i]), 'freq': to_int(f.getlist('t_freq')[i]),
                              'workers': to_int(f.getlist('t_workers')[i]) if i < len(f.getlist('t_workers')) else None})
        a.tasks = jdump(tasks)
        a.checklist = jdump([{'unit': t['name'], 'items': t['items']} for t in tasks])   # 체크리스트 = 단위작업별 호
        items_of = {t['name']: t['items'] for t in tasks}
        # 3단계: 단위작업명 | 부담작업(호: 자동) | 유해요인 | 발생원인 | 비고
        hz = []
        for tk, ty, ca, nt in zip(f.getlist('h_task'), f.getlist('h_type'), f.getlist('h_cause'), f.getlist('h_note')):
            if ca.strip() or ty:
                hz.append({'task': tk.strip(), 'items': items_of.get(tk.strip(), []), 'type': ty, 'cause': ca.strip(), 'note': nt.strip()})
        a.hazards = jdump(hz)
        # 3-1: 인간공학적 작업·분석평가 (선택)
        eg = []
        for un, tl, rs_, jd in zip(f.getlist('e_unit'), f.getlist('e_tool'), f.getlist('e_result'), f.getlist('e_judge')):
            if un.strip() or rs_.strip():
                eg.append({'unit': un.strip(), 'tool': tl, 'result': rs_.strip(), 'judge': jd.strip()})
        a.ergo = jdump(eg)
        a.summary = f.get('summary', '').strip()
        a.note = f.get('note', '').strip()
        db.session.flush()
        for up in request.files.getlist('photos'):
            if up and up.filename:
                db.session.add(Photo(assessment_id=a.id, caption=f.get('photo_caption', '').strip() or up.filename,
                                     data=up.read(), mime=up.mimetype))
        db.session.commit()
        flash('유해요인조사표를 저장했습니다.')
        return redirect(url_for('assess', jid=jid))
    photos = Photo.query.filter_by(assessment_id=a.id).with_entities(Photo.id, Photo.caption).all() if a else []
    rs = Response.query.filter_by(round_id=rnd.id, job_id=jid).all()
    workers = Employee.query.filter_by(job_id=jid).order_by(Employee.name).all()
    tasks = a.task_list if a and a.task_list else [{'name': u['name'], 'items': u.get('items', [])} for u in job.unit_list]
    return render_template('admin_assess.html', job=job, a=a, rnd=rnd, photos=photos, rs=rs, workers=workers,
                           ustats=unit_stats(rnd.id, jid), udesc={u['name']: u for u in job.unit_list},
                           st=jload(a.site_status) if a else {}, tasks=tasks,
                           hazards=jload(a.hazards, []) if a else [], ergo=jload(a.ergo, []) if a else [])


@app.route('/admin/photo/<int:pid>')
@admin_required
def photo(pid):
    p = db.session.get(Photo, pid) or abort(404)
    return HttpResponse(p.data, mimetype=p.mime or 'image/jpeg')


@app.route('/admin/photo/<int:pid>/delete', methods=['POST'])
@admin_required
def photo_delete(pid):
    p = db.session.get(Photo, pid) or abort(404)
    a = db.session.get(Assessment, p.assessment_id)
    db.session.delete(p)
    db.session.commit()
    if request.headers.get('X-Requested-With') == 'fetch':
        return {'ok': True}
    return redirect(url_for('assess', jid=a.job_id))


@app.route('/admin/assess/<int:jid>/photo', methods=['POST'])
@admin_required
def photo_upload(jid):
    """아이패드 현장 촬영 사진 즉시 업로드(조사표 저장 전이라도 보관). 없으면 조사표를 먼저 만든다"""
    rnd = current_round() or abort(400)
    db.session.get(Job, jid) or abort(404)
    a = Assessment.query.filter_by(round_id=rnd.id, job_id=jid).first()
    if not a:
        a = Assessment(round_id=rnd.id, job_id=jid)
        db.session.add(a)
        db.session.flush()
    up = request.files.get('photo')
    if not up:
        abort(400)
    p = Photo(assessment_id=a.id, caption=(request.form.get('caption') or '').strip() or '작업 사진',
              data=up.read(), mime=up.mimetype or 'image/jpeg')
    db.session.add(p)
    db.session.commit()
    return {'id': p.id, 'url': url_for('photo', pid=p.id), 'caption': p.caption,
            'delete': url_for('photo_delete', pid=p.id)}


# ── 응답 관리 ─────────────────────────────────────────────────────────
@app.route('/admin/responses')
@admin_required
def responses():
    rnd = current_round()
    q = Response.query.filter_by(round_id=rnd.id) if rnd else Response.query.filter(False)
    dept = request.args.get('dept')
    level = request.args.get('level')
    if dept:
        q = q.filter_by(dept=dept)
    if level == 'pain':
        q = q.filter(Response.top_level == 'pain')
    elif level == 'manage':
        q = q.filter(Response.top_level.in_(['manage', 'pain']))
    elif level == 'complaint':
        q = q.filter(Response.top_level.isnot(None))
    elif level == 'none':
        q = q.filter(Response.top_level.is_(None))
    rows = q.order_by(Response.id.desc()).all()
    depts = sorted({r.dept for r in Response.query.filter_by(round_id=rnd.id)} if rnd else [])
    # 미응답자: 직원명단 중 이 회차에 응답(employee_id) 없는 사람 (휴직·휴가 제외)
    done_ids = {r.employee_id for r in Response.query.filter_by(round_id=rnd.id)} if rnd else set()
    pending = [e for e in Employee.query.order_by(Employee.dept, Employee.name) if e.id not in done_ids and not e.on_leave]
    leave = [e for e in Employee.query if e.on_leave]
    return render_template('admin_responses.html', rows=rows, depts=depts, dept=dept, level=level, pending=pending, leave=leave,
                           total_emp=Employee.query.count())


@app.route('/admin/response/<int:rid>', methods=['GET', 'POST'])
@admin_required
def response_detail(rid):
    r = db.session.get(Response, rid) or abort(404)
    if request.method == 'POST':
        r.followup = request.form.get('followup')
        r.admin_note = request.form.get('admin_note', '').strip()
        jid = to_int(request.form.get('job_id'))
        if request.form.get('job_id') is not None:
            r.job_id = jid or None
        db.session.commit()
        flash('저장했습니다.')
        return redirect(url_for('response_detail', rid=rid))
    jobs_all = Job.query.order_by(Job.dept, Job.name).all()
    return render_template('admin_response.html', r=r, jobs_all=jobs_all, jload=jload)


@app.route('/admin/response/<int:rid>/delete', methods=['POST'])
@admin_required
def response_delete(rid):
    r = db.session.get(Response, rid) or abort(404)
    db.session.delete(r)
    db.session.commit()
    flash('삭제했습니다.')
    return redirect(url_for('responses'))


# ── 개선계획서 ────────────────────────────────────────────────────────
@app.route('/admin/improvements', methods=['GET', 'POST'])
@admin_required
def improvements():
    rnd = current_round() or abort(400)
    if request.method == 'POST':
        f = request.form
        iid = to_int(f.get('id'))
        m = db.session.get(Improvement, iid) if iid else Improvement(round_id=rnd.id)
        m.job_id = to_int(f.get('job_id')) or None
        for k in ('problem', 'opinion', 'plan', 'schedule', 'status', 'result'):
            setattr(m, k, (f.get(k) or '').strip())
        m.cost = to_int(f.get('cost'))
        m.priority = to_int(f.get('priority'))
        m.done_date = to_date(f.get('done_date'))
        if not iid:
            db.session.add(m)
        db.session.commit()
        flash('저장했습니다.')
        return redirect(url_for('improvements'))
    A = analyze(rnd)
    edit = db.session.get(Improvement, to_int(request.args.get('edit'))) if request.args.get('edit') else None
    prefill = None
    if not edit and request.args.get('job'):
        jid = to_int(request.args.get('job'))
        a = Assessment.query.filter_by(round_id=rnd.id, job_id=jid).first()
        hz = jload(a.hazards, []) if a else []
        ops = [r.opinion_ko for r in Response.query.filter_by(round_id=rnd.id, job_id=jid) if r.opinion_ko]
        rank = next((x['rank'] for x in A['ranked'] if x['job'].id == jid), None)
        prefill = {'job_id': jid, 'priority': rank,
                   'problem': '\n'.join([f"[종합판정] {jx['q_label']} — {jx['why']}" for jx in A['job_summary'] if jx['job'].id == jid][:1]
                                        + [f"[{h.get('type')}] {h.get('cause')}" for h in hz if h.get('cause')]),
                   'opinion': '\n'.join(ops)}
    jobs_all = Job.query.order_by(Job.dept, Job.name).all()
    return render_template('admin_improvements.html', A=A, rows=A['improvements'], edit=edit,
                           prefill=prefill, jobs_all=jobs_all)


@app.route('/admin/improvement/<int:iid>/delete', methods=['POST'])
@admin_required
def improvement_delete(iid):
    m = db.session.get(Improvement, iid) or abort(404)
    db.session.delete(m)
    db.session.commit()
    return redirect(url_for('improvements'))


# ── 보고서 ────────────────────────────────────────────────────────────
def _download(data, filename, mime):
    return HttpResponse(data, mimetype=mime, headers={
        'Content-Disposition': f"attachment; filename*=UTF-8''{quote(filename)}"})


@app.route('/admin/report')
@admin_required
def report():
    rnd = current_round()
    if not rnd:
        return redirect(url_for('rounds'))
    A = analyze(rnd)
    missing = [j for j in A['job_summary'] if j['job'].active and not j['assess']]
    return render_template('admin_report.html', A=A, rnd=rnd, missing=missing)


@app.route('/admin/report.xlsx')
@admin_required
def report_xlsx():
    import reports
    rnd = current_round() or abort(400)
    data = reports.build_excel(analyze(rnd), COMPANY, SITE)
    stamp = kst(datetime.utcnow()).strftime('%Y%m%d')
    return _download(data, f'{rnd.title}_결과보고서_{stamp}.xlsx',
                     'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/admin/report.pdf')
@admin_required
def report_pdf():
    import reports
    rnd = current_round() or abort(400)
    photos = {}
    for a in Assessment.query.filter_by(round_id=rnd.id):
        photos[a.id] = [(p.caption, p.data) for p in Photo.query.filter_by(assessment_id=a.id)]
    data = reports.build_pdf(analyze(rnd), COMPANY, SITE, photos)
    stamp = kst(datetime.utcnow()).strftime('%Y%m%d')
    return _download(data, f'{rnd.title}_결과보고서_{stamp}.pdf', 'application/pdf')


@app.route('/admin/responses.xlsx')
@admin_required
def responses_xlsx():
    import reports
    rnd = current_round() or abort(400)
    data = reports.build_raw_excel(analyze(rnd))
    stamp = kst(datetime.utcnow()).strftime('%Y%m%d')
    return _download(data, f'{rnd.title}_증상조사원자료_{stamp}.xlsx',
                     'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/admin/reload-texts')
@admin_required
def reload_texts():
    load_texts()
    flash('번역 사전(i18n/survey_texts.json)을 다시 읽었습니다.')
    return redirect(request.referrer or url_for('admin'))


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5003)), debug=True)
