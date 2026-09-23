# -*- coding: utf-8 -*-
"""증상조사표 다국어 사전 생성 (무과금 deep_translator)

    python make_i18n.py            # 없는 키만 번역해서 추가 (검토·수정한 번역은 보존)
    python make_i18n.py --force    # 전부 다시 번역

결과: i18n/survey_texts.json  →  사람이 직접 열어 번역을 고쳐도 됨(앱이 그대로 읽음).

구글 웹 번역은 호출이 잦으면 PC(IP)를 한동안 차단한다. 그래서 문구 여러 개를
'#번호# 문장' 줄로 묶어 한 번에 보내고(언어당 1~2회), 파싱 못 한 것만 개별 호출한다.
"""
import json, os, re, sys, time
from deep_translator import GoogleTranslator
from survey_def import KO, LANGS

BASE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(BASE, 'i18n', 'survey_texts.json')
CHUNK = 3500   # 구글 1회 5000자 제한


class Blocked(Exception):
    pass


def call(text, lang):
    for attempt in range(2):
        try:
            time.sleep(1.5)
            return GoogleTranslator(source='ko', target=lang).translate(text)
        except Exception as e:
            if 'TooManyRequests' in type(e).__name__:
                raise Blocked()          # 차단 중엔 더 호출하면 차단이 길어짐 → 즉시 중단
            print(f'  retry {attempt + 1}: {type(e).__name__}')
            time.sleep(5)
    return None


def batch(keys, lang):
    out, buf = {}, []

    def flush():
        if not buf:
            return
        src = '\n'.join(f'#{i}# {KO[k]}' for i, k in buf)
        res = call(src, lang) or ''
        idx = dict(buf)
        for m in re.finditer(r'#\s*(\d+)\s*#\s*(.+?)(?=\s*#\s*\d+\s*#|\Z)', res, re.S):
            i = int(m.group(1))
            if i in idx and m.group(2).strip():
                out[idx[i]] = m.group(2).strip()
        buf.clear()

    size = 0
    for i, k in enumerate(keys):
        ln = len(KO[k]) + 8
        if size + ln > CHUNK:
            flush()
            size = 0
        buf.append((i, k))
        size += ln
    flush()
    for k in keys:                # 묶음에서 빠진 것만 개별 번역
        if k not in out:
            v = call(KO[k], lang)
            if v:
                out[k] = v
    return out


def main():
    force = '--force' in sys.argv
    data = {}
    if os.path.exists(OUT) and not force:
        with open(OUT, encoding='utf-8') as f:
            data = json.load(f)
    only = [a for a in sys.argv[1:] if a in LANGS]
    for lang in (only or LANGS):
        if lang == 'ko':
            continue
        cur = data.setdefault(lang, {})
        todo = [k for k in KO if force or not cur.get(k)]
        print(f'{lang}: {len(todo)}개 번역')
        if todo:
            try:
                cur.update(batch(todo, lang))
            except Blocked:
                print('  ✗ 구글 번역이 이 PC를 일시 차단 중 → 몇 시간 뒤 다시 실행하세요. (지금까지 번역분은 저장)')
                break
            print(f'  → {sum(1 for k in KO if cur.get(k))}/{len(KO)}')
        os.makedirs(os.path.dirname(OUT), exist_ok=True)
        with open(OUT, 'w', encoding='utf-8') as f:     # 언어마다 저장 (중간에 끊겨도 보존)
            json.dump(data, f, ensure_ascii=False, indent=1)
    print('saved', OUT)
    return blocked_any(data)


def blocked_any(data):
    return not all(data.get(lg, {}).get(k) for lg in LANGS if lg != 'ko' for k in KO)


def translate_job_units():
    """작업 마스터의 단위작업(이름·설명) 번역 캐시(Job.units_tr) 채우기 — 앱 DB(msd.db) 직접 갱신"""
    import app as A
    with A.app.app_context():
        todo = [j for j in A.Job.query if j.unit_list and not j.units_tr]
        print(f'단위작업 번역 대상 작업: {len(todo)}개')
        for j in todo:
            tr = A.translate_units(j.unit_list)
            if len(tr) < len([l for l in LANGS if l != 'ko']):
                print('  ✗ 구글 차단/실패 — 단위작업 번역 중단 (나중에 다시 실행)')
                break
            j.units_tr = json.dumps(tr, ensure_ascii=False)
            A.db.session.commit()
            print(f'  ✓ {j.dept} / {j.name}')
            time.sleep(2)


if __name__ == '__main__':
    incomplete = main()
    if incomplete:
        print('문구 번역이 아직 덜 끝나(차단 등) 단위작업 번역은 건너뜀')
    elif '--skip-units' not in sys.argv:
        translate_job_units()
